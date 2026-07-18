
import * as _fs_hooks from 'fs';
const _originalLog = console.log;
console.log = function(...args) {
  _originalLog.apply(console, args);
  try { _fs_hooks.appendFileSync('sync_debug.log', args.join(' ') + '\n'); } catch (e) {}
};


/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import express from "express";
import fs from "fs";
import path from "path";
import dotenv from "dotenv";
dotenv.config();
import { createServer as createViteServer } from "vite";
import { google } from "googleapis";
import { 
  loadDB, saveDB, runSystemScans, initFirestoreDB, hasPendingChanges, setHasPendingChanges
} from "./src/server/db.js";
import { 
  UserRole, ServantStatus, Servant, Location, DistributionItem, 
  Attendance, Leave, Suspension, ExcludedServant, Message, 
  AuditLog, SystemNotification, SheetsSyncLog, Duty, FinanceCampaign, FinanceRecord
} from "./src/types.js";


function resolveSpreadsheetId(customId?: string) {
  const db = loadDB();
  return (customId || db.sheetsSyncConfig?.spreadsheetId || "").trim();
}

const app = express();
let isDbReady = false;
let dbReadyPromise = null;

// Middleware to block API requests until DB is ready
app.use("/api", async (req, res, next) => {
  if (isDbReady) {
    return next();
  }
  if (!dbReadyPromise) {
    return res.status(503).json({ error: "الخادم قيد الإعداد وجلب البيانات من Google Sheets. يرجى الانتظار..." });
  }
  try {
    await dbReadyPromise;
    next();
  } catch (e) {
    next();
  }
});

const PORT = 3000;

app.use(express.json({ limit: "10mb" }));
app.use(express.urlencoded({ limit: "10mb", extended: true }));

// API Middlewares / Helper to log audit actions
function auditLog(username: string, actionType: "ADD" | "EDIT" | "DELETE" | "SYNC" | "DISTRIBUTE" | "AUTH", details: string) {
  const db = loadDB();
  const todayStr = new Date().toISOString().split("T")[0];
  const timeStr = new Date().toTimeString().split(" ")[0];
  
  const log: AuditLog = {
    id: `log-${Date.now()}-${Math.random().toString(36).substr(2, 4)}`,
    username,
    date: todayStr,
    time: timeStr,
    actionType,
    details
  };
  
  db.auditLogs.unshift(log); // Keep latest at top
  saveDB(db).catch(console.error);
}

function getBaghdadDate(): Date {
  const now = new Date();
  const formatter = new Intl.DateTimeFormat("en-US", {
    timeZone: "Asia/Baghdad",
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  });
  const parts = formatter.formatToParts(now);
  const year = parseInt(parts.find(p => p.type === "year")?.value || "2026");
  const month = parseInt(parts.find(p => p.type === "month")?.value || "1") - 1;
  const day = parseInt(parts.find(p => p.type === "day")?.value || "1");
  return new Date(year, month, day);
}

function parseSheetDate(dateStr: string): Date | null {
  if (!dateStr) return null;
  const cleaned = String(dateStr).trim();
  if (!cleaned) return null;

  // Handle standard YYYY-MM-DD
  if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(cleaned)) {
    const [y, m, d] = cleaned.split("-").map(Number);
    return new Date(y, m - 1, d);
  }
  // Handle DD/MM/YYYY or MM/DD/YYYY
  if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(cleaned)) {
    const parts = cleaned.split("/").map(Number);
    const d = parts[0];
    const m = parts[1];
    const y = parts[2];
    return new Date(y, m - 1, d);
  }
  // Fallback
  const d = new Date(cleaned);
  if (isNaN(d.getTime())) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

// Helper to handle retry of Google Sheets API calls in case of Rate Limit / Quota Exceeded error (HTTP 429)
async function withSheetsRetry<T>(fn: () => Promise<T>, retries = 5, delay = 1000): Promise<T> {
  try {
    return await fn();
  } catch (error: any) {
    const isRateLimit = error.status === 429 || 
                       error.code === 429 || 
                       String(error.message || "").toLowerCase().includes("quota exceeded") || 
                       String(error.message || "").toLowerCase().includes("rate limit");
    if (isRateLimit && retries > 0) {
      console.warn(`[Sheets Retry] Rate limit hit (${error.message || error.status}). Retrying in ${delay}ms... (Retries left: ${retries})`);
      await new Promise(resolve => setTimeout(resolve, delay));
      return withSheetsRetry(fn, retries - 1, delay * 1.5); // Exponential backoff
    }
    throw error;
  }
}

// Helper to sanitize cell values to prevent the 50,000 character Google Sheets cell limit error
function sanitizeCellForGoogleSheets(val: any): string {
  if (val && typeof val === 'string' && val.startsWith('data:image')) {
    console.log('[PHOTO_PUSHED] Google Sheets (attempting to push photo of length: ' + val.length + ')');
  }
  if (val === null || val === undefined) return "";
  const str = String(val);
  if (str.length > 48000) {
    if (str.startsWith("data:image") || str.includes(";base64,")) {
      console.log('[PHOTO_PUSHED] Google Sheets (photo too large, replacing with placeholder: ' + str.length + ')');
      return "[صورة شخصية مرفوعة من النظام - حجمها كبير جداً]";
    }
    if (str.startsWith("data:image") || str.includes(";base64,")) {
      return "[صورة شخصية مرفوعة من النظام - حجمها كبير جداً]";
    }
    return str.substring(0, 48000) + "... [تم تقليص النص لتجاوز الحد الأقصى]";
  }
  return str;
}

// Helper to merge comma-separated historical lists in "السجل الدائم"
function mergeCommaValues(existingVal: string | null | undefined, newVal: string | null | undefined): string {
  const valSet = new Set<string>();
  
  if (existingVal) {
    String(existingVal).split(",").forEach(v => {
      const trimmed = v.trim();
      if (trimmed && trimmed !== "بلا" && trimmed !== "-") valSet.add(trimmed);
    });
  }
  
  if (newVal) {
    String(newVal).split(",").forEach(v => {
      const trimmed = v.trim();
      if (trimmed && trimmed !== "بلا" && trimmed !== "-") valSet.add(trimmed);
    });
  }
  
  if (valSet.size === 0) return "";
  return Array.from(valSet).join(", ");
}

// Helper to merge list sheets (سجل الاجازات, الغيابات, الموقوفين, المستبعدين) to preserve deleted records
function mergeSheetRows(existingRows: any[][] | null | undefined, newRows: any[][], keyIndices: number[]): any[][] {
  if (!existingRows || existingRows.length <= 1) {
    return newRows;
  }
  
  const merged: any[][] = [...existingRows];
  
  // Create a set of keys for existing rows so we don't duplicate
  const existingKeys = new Set<string>();
  for (let i = 1; i < existingRows.length; i++) {
    const row = existingRows[i];
    if (!row || row.length === 0) continue;
    
    // Construct unique key based on the key indices
    const key = keyIndices.map(idx => String(row[idx] || "").trim()).join("|");
    if (key) {
      existingKeys.add(key);
    }
  }
  
  // For each of the new rows (excluding header), if its key doesn't exist, append it!
  for (let i = 1; i < newRows.length; i++) {
    const row = newRows[i];
    if (!row || row.length === 0) continue;
    
    const key = keyIndices.map(idx => String(row[idx] || "").trim()).join("|");
    if (!existingKeys.has(key)) {
      merged.push(row);
    } else {
      // If key exists, let's update the existing row with new values
      const idx = merged.findIndex((r, index) => {
        if (index === 0) return false;
        const rKey = keyIndices.map(idx => String(r[idx] || "").trim()).join("|");
        return rKey === key;
      });
      if (idx !== -1) {
        merged[idx] = row;
      }
    }
  }
  
  return merged;
}



// -----------------------------------------------------------------


// -----------------------------------------------------------------
// AUTH ENDPOINTS
// -----------------------------------------------------------------
app.post("/api/auth/login", async (req, res) => {
  const { code } = req.body;
  const db = loadDB();

  if (!code) {
    return res.status(400).json({ success: false, message: "الرجاء إدخال كود الدخول" });
  }

  const cleanInputCode = code.trim();

  // 1. Admin login check
  const isAdminCode = cleanInputCode === "1972" || cleanInputCode === "197200";

  if (isAdminCode) {
    const user = {
      id: "u-admin",
      username: cleanInputCode,
      name: "محمد زيدان عباس",
      role: UserRole.ADMIN,
      canPushData: true
    };
    auditLog("محمد زيدان عباس", "AUTH", "تسجيل دخول ناجح للمدير العام باستخدام الكود فقط");
    return res.json({ success: true, user });
  }

  // 2. Servant / Supervisor login check: Strictly match by servant code (الكود) only
  const servant = db.servants.find(s => 
    s.code && String(s.code).trim() === cleanInputCode
  );
  if (servant) {
    const role = servant.isSupervisor ? UserRole.SUPERVISOR : UserRole.SERVANT;
    const user = {
      id: servant.id,
      username: servant.code,
      name: servant.name,
      role: role,
      code: servant.code, // Guaranteed internal code
      canPushData: servant.canPushData
    };
    
    const roleLabel = role === UserRole.SUPERVISOR ? "المشرف" : "الخادم";
    auditLog(servant.name, "AUTH", `تسجيل دخول ناجح لـ ${roleLabel} ${servant.name} باستخدام الكود الخاص به`);
    return res.json({ success: true, user, servant });
  }

  return res.status(404).json({ success: false, message: "رقم الكود غير مسجل بنظام الهيئة" });
});

// -----------------------------------------------------------------
// ACTIVE USERS (PRESENCE) TRACKING
// -----------------------------------------------------------------
interface OnlineUser {
  id: string;
  name: string;
  role: string;
  lastActive: number;
}

const onlineUsers: Record<string, OnlineUser> = {};

// Clean up inactive users (no ping in the last 45 seconds) every 15 seconds
setInterval(() => {
  const now = Date.now();
  for (const [id, user] of Object.entries(onlineUsers)) {
    if (now - user.lastActive > 45000) {
      delete onlineUsers[id];
    }
  }
}, 15000);

// Endpoint to register a ping
app.post("/api/users/ping", async (req, res) => {
  const { id, name, role } = req.body;
  if (id && name) {
    onlineUsers[id] = {
      id,
      name,
      role: role || "SERVANT",
      lastActive: Date.now()
    };
  }
  res.json({ success: true });
});

// Endpoint to get the list of active users
app.get("/api/users/online", async (req, res) => {
  const now = Date.now();
  const activeList = Object.values(onlineUsers).filter(
    (u) => now - u.lastActive <= 45000
  );
  res.json({
    success: true,
    count: activeList.length,
    users: activeList.map(u => ({ id: u.id, name: u.name, role: u.role }))
  });
});

// -----------------------------------------------------------------
// SERVANTS ENDPOINTS
// -----------------------------------------------------------------
app.get("/api/servants", async (req, res) => {
  const db = loadDB();
  res.json(db.servants);
});

// New endpoint for admin to get all members and their barcodes (access codes)
app.get("/api/admin/servants-barcodes", async (req, res) => {
  const db = loadDB();
  const data = db.servants.map(s => {
    return {
      id: s.id,
      name: s.name,
      code: s.code,
      accessCode: s.code, // Set to servant's code
      phone: s.phone || "",
      birthDate: s.birthDate || "",
      avatar: s.avatar || "",
      photoUrl: s.photoUrl || ""
    };
  });
  res.json({ success: true, data });
});

app.post("/api/servants", async (req, res) => {
  let { code, name, phone, joinDate, notes, status, humanitarian, humanitarianReason, isSupervisor, birthDate, address, avatar, accessCode } = req.body;
  const db = loadDB();

  if (!code || !name) {
    return res.status(400).json({ error: "الكود والاسم الرباعي مطلوبان" });
  }

  // Code unique check
  if (db.servants.some(s => s.code === code)) {
    return res.status(400).json({ error: "رقم الكود مسجل بالفعل لخادم آخر" });
  }

  // We map accessCode directly to code since separate accessCode is deprecated
  accessCode = code;

  const newServant: Servant = {
    id: `s-${code}`,
    code,
    name,
    phone,
    joinDate: joinDate || new Date().toISOString().split("T")[0],
    notes,
    status: status || ServantStatus.ACTIVE,
    statusDate: new Date().toISOString().split("T")[0],
    humanitarian: !!humanitarian,
    humanitarianReason: humanitarian ? humanitarianReason : "",
    isSupervisor: !!isSupervisor,
    birthDate: birthDate || "",
    address: address || "",
    avatar: avatar || "",
    accessCode
  };

  db.servants.push(newServant);
  db.servants.sort((a, b) => {
    const codeA = parseInt(a.code) || 0;
    const codeB = parseInt(b.code) || 0;
    return codeA - codeB;
  });
  
  // Add notification
  db.notifications.unshift({
    id: `notif-${Date.now()}`,
    title: "إضافة خادم جديد",
    content: `تم تسجيل الخادم الجديد ${name} بكود (${code}) في الهيئة بنجاح.`,
    type: "SUCCESS",
    date: `${new Date().toISOString().split("T")[0]} 12:00`,
    isRead: false
  });

  saveDB(db).catch(console.error);
  auditLog("admin", "ADD", `إضافة الخادم الجديد: ${name} (كود: ${code})`);
  pushAllServantsToGoogleSheet(true).catch(console.error);
  res.status(201).json(newServant);
});

app.put("/api/servants/:id", async (req, res) => {
  if (req.body.avatar && req.body.avatar.startsWith('data:image')) {
    console.log('[PHOTO_UPLOAD] Servant=' + req.body.name + ', originalLength=' + req.body.avatar.length);
  }
  const { id } = req.params;
  const { name, phone, notes, status, humanitarian, humanitarianReason, pinnedLocationId, isSupervisor, birthDate, address, avatar, accessCode } = req.body;
  const db = loadDB();

  const idx = db.servants.findIndex(s => s.id === id);
  if (idx === -1) return res.status(404).json({ error: "الخادم غير موجود" });

  const oldServant = db.servants[idx];
  const oldStatus = oldServant.status;

  const finalStatus = status || oldServant.status;
  const finalStatusDate = finalStatus !== oldStatus ? new Date().toISOString().split("T")[0] : (oldServant.statusDate || oldServant.joinDate || new Date().toISOString().split("T")[0]);

  // Update
  db.servants[idx] = {
    ...oldServant,
    name: name || oldServant.name,
    phone: phone !== undefined ? phone : oldServant.phone,
    notes: notes !== undefined ? notes : oldServant.notes,
    status: finalStatus,
    statusDate: finalStatusDate,
    humanitarian: humanitarian !== undefined ? !!humanitarian : oldServant.humanitarian,
    humanitarianReason: humanitarian ? humanitarianReason : "",
    pinnedLocationId: pinnedLocationId !== undefined ? pinnedLocationId : oldServant.pinnedLocationId,
    isSupervisor: isSupervisor !== undefined ? !!isSupervisor : oldServant.isSupervisor,
    birthDate: birthDate !== undefined ? birthDate : oldServant.birthDate,
    address: address !== undefined ? address : oldServant.address,
    avatar: avatar !== undefined ? avatar : oldServant.avatar,
    accessCode: oldServant.code
  };

  // If status changed to EXCLUDED, log it & add to excluded register
  if (status === ServantStatus.EXCLUDED && oldStatus !== ServantStatus.EXCLUDED) {
    const exists = db.excluded.some(e => e.servantId === id);
    if (!exists) {
      db.excluded.push({
        id: `excl-${Date.now()}`,
        servantId: id,
        servantCode: oldServant.code,
        servantName: oldServant.name,
        reason: notes || "تم الاستبعاد بناءً على قرار إداري",
        date: new Date().toISOString().split("T")[0],
        notes: "تم نقله تلقائياً لسجل المستبعدين"
      });
    }
  }

  // Add notification about edit
  db.notifications.unshift({
    id: `notif-${Date.now()}`,
    title: "تحديث بيانات خادم",
    content: `تم تعديل بيانات الخادم ${oldServant.name} (كود: ${oldServant.code}) بنجاح.`,
    type: "INFO",
    date: `${new Date().toISOString().split("T")[0]} 12:00`,
    isRead: false
  });

  saveDB(db).catch(console.error);
  auditLog("admin", "EDIT", `تعديل بيانات الخادم: ${oldServant.name} (كود: ${oldServant.code})`);
  pushAllServantsToGoogleSheet(true).catch(console.error);
  res.json(db.servants[idx]);
});

app.delete("/api/servants/:id", async (req, res) => {
  const { id } = req.params;
  const db = loadDB();

  const idx = db.servants.findIndex(s => s.id === id);
  if (idx === -1) return res.status(404).json({ error: "الخادم غير موجود" });

  const servant = db.servants[idx];
  db.servants.splice(idx, 1);

  // Clean related records in other tables
  db.distributions = db.distributions.filter(d => d.servantId !== id);
  db.attendance = db.attendance.filter(a => a.servantId !== id);
  db.leaves = db.leaves.filter(l => l.servantId !== id);
  db.suspensions = db.suspensions.filter(s => s.servantId !== id);
  db.excluded = db.excluded.filter(e => e.servantId !== id);

  // Notification
  db.notifications.unshift({
    id: `notif-${Date.now()}`,
    title: "حذف خادم من الهيئة",
    content: `تم إزالة ملف الخادم ${servant.name} (كود: ${servant.code}) وسجلاته نهائياً من نظام الهيئة.`,
    type: "WARNING",
    date: `${new Date().toISOString().split("T")[0]} 12:00`,
    isRead: false
  });

  saveDB(db).catch(console.error);
  auditLog("admin", "DELETE", `حذف ملف الخادم: ${servant.name} (كود: ${servant.code})`);
  pushAllServantsToGoogleSheet(true).catch(console.error);
  res.json({ success: true, message: "تم حذف الخادم بنجاح" });
});

app.post("/api/servants/delete-all", async (req, res) => {
  const db = loadDB();
  db.servants = [];
  db.distributions = [];
  db.attendance = [];
  db.leaves = [];
  db.suspensions = [];
  db.excluded = [];

  db.notifications.unshift({
    id: `notif-${Date.now()}`,
    title: "حذف جميع الخدام وأعضاء الهيئة",
    content: "تم تفريغ وحذف سجل كافة الخدام وأعضاء الهيئة المسجلين في النظام وسجلات حضورهم والتزامهم نهائياً.",
    type: "WARNING",
    date: `${new Date().toISOString().split("T")[0]} 12:00`,
    isRead: false
  });

  saveDB(db).catch(console.error);
  auditLog("admin", "DELETE", "حذف كافة سجلات وملفات الخدام وأعضاء الهيئة وسجلاتهم نهائياً");
  pushAllServantsToGoogleSheet(true).catch(console.error);
  res.json({ success: true, message: "تم حذف جميع الخدام بنجاح" });
});

app.post("/api/servants/bulk-import", async (req, res) => {
  const { servants: importedServants } = req.body;
  if (!importedServants || !Array.isArray(importedServants)) {
    return res.status(400).json({ error: "البيانات المرسلة غير صالحة" });
  }

  const db = loadDB();
  let addedCount = 0;
  let skippedCount = 0;

  for (const item of importedServants) {
    const { code, name, phone, joinDate, notes, status, humanitarian, humanitarianReason, isSupervisor, birthDate, address } = item;
    
    // basic validations
    if (!code || !name) {
      skippedCount++;
      continue;
    }

    const cleanCode = String(code).trim();
    const cleanName = String(name).trim();

    // check duplicate
    if (db.servants.some(s => s.code === cleanCode)) {
      skippedCount++;
      continue;
    }

    let accessCode = item.accessCode ? String(item.accessCode).trim() : "";
    if (!accessCode) {
      let hash = 0;
      const codeStr = String(cleanCode);
      for (let j = 0; j < codeStr.length; j++) {
        hash = (hash << 5) - hash + codeStr.charCodeAt(j);
        hash |= 0;
      }
      accessCode = String(Math.abs(hash) % 900000 + 100000);
    }

    const newServant: Servant = {
      id: `s-${cleanCode}`,
      code: cleanCode,
      name: cleanName,
      phone: phone ? String(phone).trim() : "",
      joinDate: joinDate || new Date().toISOString().split("T")[0],
      notes: notes || "",
      status: ServantStatus.ACTIVE, // Force always ACTIVE upon import as requested
      humanitarian: !!humanitarian,
      humanitarianReason: humanitarian ? (humanitarianReason || "") : "",
      isSupervisor: !!isSupervisor,
      birthDate: birthDate ? String(birthDate).trim() : "",
      address: address ? String(address).trim() : "",
      accessCode
    };

    db.servants.push(newServant);
    addedCount++;
    
    // Real-time sync push to Google Sheets (non-blocking)
    
  }

  if (addedCount > 0) {
    db.notifications.unshift({
      id: `notif-${Date.now()}`,
      title: "رفع واستيراد جماعي للخدام",
      content: `تم استيراد عدد (${addedCount}) خادم بنجاح إلى سجلات الهيئة من ملف اكسل. تم تخطي (${skippedCount}) سجل بسبب التكرار أو نقص البيانات.`,
      type: "SUCCESS",
      date: `${new Date().toISOString().split("T")[0]} 12:00`,
      isRead: false
    });

    db.servants.sort((a, b) => {
      const codeA = parseInt(a.code) || 0;
      const codeB = parseInt(b.code) || 0;
      return codeA - codeB;
    });
    saveDB(db).catch(console.error);
    auditLog("admin", "ADD", `استيراد جماعي للخدام: تم إضافة ${addedCount} خادم وتخطي ${skippedCount}`);
    pushAllServantsToGoogleSheet(true).catch(console.error);
  }
  res.json({ success: true, addedCount, skippedCount, servants: db.servants });
});

// -----------------------------------------------------------------
// DUTIES ENDPOINTS
// -----------------------------------------------------------------
app.get("/api/duties", async (req, res) => {
  const db = loadDB();
  res.json(db.duties);
});

app.post("/api/duties", async (req, res) => {
  const { date, label } = req.body;
  const db = loadDB();

  if (!date || !label) {
    return res.status(400).json({ error: "التاريخ وتسمية المناسبة حقول إلزامية" });
  }

  const dateStr = date.trim();
  const exists = db.duties.find(d => d.date === dateStr);
  if (exists) {
    return res.status(400).json({ error: "يوجد واجب مسجل بالفعل في هذا التاريخ" });
  }

  const dObj = new Date(dateStr);
  const daysMap: Record<number, string> = {
    0: "Sunday",
    1: "Monday",
    2: "Tuesday",
    3: "Wednesday",
    4: "Thursday",
    5: "Friday",
    6: "Saturday"
  };
  const dayOfWeek = daysMap[dObj.getDay()] || "Special";

  const newDuty: Duty = {
    id: `duty-special-${Date.now()}`,
    date: dateStr,
    dayOfWeek: dayOfWeek,
    year: dObj.getFullYear() || 2026,
    label: label.trim()
  };

  db.duties.push(newDuty);
  
  // Create system notification for a new special occasion
  db.notifications.push({
    id: `notif-${Date.now()}-special-duty`,
    title: `مناسبة خاصة جديدة: ${newDuty.label}`,
    content: `تم إضافة يوم إضافي كمناسبة خاصة بتاريخ ${newDuty.date} وتخصيصها لتوزيع كشوفات الخدمة.`,
    type: "INFO",
    date: `${new Date().toISOString().split("T")[0]} 12:00`,
    isRead: false
  });

  saveDB(db).catch(console.error);
  auditLog("admin", "ADD", `إضافة واجب مناسبة خاصة: ${newDuty.label} في تاريخ ${newDuty.date}`);

  res.json({ success: true, duty: newDuty });
});

app.put("/api/duties/:id", async (req, res) => {
  const { id } = req.params;
  const { date, label } = req.body;
  const db = loadDB();
  const duty = db.duties.find(d => d.id === id);
  if (!duty) return res.status(404).json({ error: "الواجب غير موجود" });

  if (date) {
    const dateStr = date.trim();
    const dObj = new Date(dateStr);
    const daysMap: Record<number, string> = {
      0: "Sunday", 1: "Monday", 2: "Tuesday", 3: "Wednesday", 4: "Thursday", 5: "Friday", 6: "Saturday"
    };
    duty.date = dateStr;
    duty.dayOfWeek = daysMap[dObj.getDay()] || "Special";
    duty.year = dObj.getFullYear() || 2026;
  }
  if (label) {
    duty.label = label.trim();
  }

  saveDB(db).catch(console.error);
  auditLog("admin", "EDIT", `تعديل بيانات واجب الخدمة: ${duty.label} (بتاريخ ${duty.date})`);
  

  res.json({ success: true, duty });
});

app.delete("/api/duties/:id", async (req, res) => {
  const { id } = req.params;
  const db = loadDB();
  const index = db.duties.findIndex(d => d.id === id);
  if (index === -1) return res.status(404).json({ error: "الواجب غير موجود" });

  const duty = db.duties[index];
  db.duties.splice(index, 1);

  // Clean up associated distributions & attendance
  db.distributions = db.distributions.filter((dist: any) => dist.dutyId !== id);
  db.attendance = db.attendance.filter((att: any) => att.dutyId !== id);

  saveDB(db).catch(console.error);
  auditLog("admin", "DELETE", `حذف واجب الخدمة: ${duty.label} (بتاريخ ${duty.date}) وكافة بيانات التوزيع والتحضير التابعة له`);
  

  res.json({ success: true });
});

app.post("/api/duties/delete-all", async (req, res) => {
  const db = loadDB();
  const count = db.duties.length;
  db.duties = [];
  db.distributions = [];
  db.attendance = [];

  saveDB(db).catch(console.error);
  auditLog("admin", "DELETE", `حذف كافة واجبات الخدمة (${count} واجب) وتفريغ كافة كشوف التوزيع والتحضير بالكامل`);
  

  res.json({ success: true, count });
});

app.post("/api/duties/:id/toggle-auto", async (req, res) => {
  const { id } = req.params;
  const db = loadDB();
  const duty = db.duties.find(d => d.id === id);
  if (!duty) return res.status(404).json({ error: "الواجب غير موجود" });

  duty.isAutoDistributeDisabled = !duty.isAutoDistributeDisabled;
  saveDB(db).catch(console.error);

  auditLog(
    "admin",
    "EDIT",
    `تعديل حالة التوزيع التلقائي لواجب ${duty.label}: ${duty.isAutoDistributeDisabled ? "معطل" : "مفعل"}`
  );

  res.json({ success: true, duty });
});

// -----------------------------------------------------------------
// FINANCE ENDPOINTS (سجل المالية للخدام)
// -----------------------------------------------------------------
app.get("/api/finance/campaigns", async (req, res) => {
  const db = loadDB();
  res.json(db.financeCampaigns || []);
});

app.post("/api/finance/campaigns", async (req, res) => {
  const { title, date, notes } = req.body;
  const db = loadDB();

  if (!title || !date) {
    return res.status(400).json({ error: "عنوان السجل والتاريخ حقول إلزامية" });
  }

  const newCampaign: FinanceCampaign = {
    id: `campaign-${Date.now()}`,
    title: title.trim(),
    date: date,
    notes: notes?.trim() || "",
    createdAt: new Date().toISOString()
  };

  db.financeCampaigns = db.financeCampaigns || [];
  db.financeCampaigns.push(newCampaign);

  // Pre-seed records for all current active/on-leave servants
  db.financeRecords = db.financeRecords || [];
  const activeServants = db.servants || [];
  
  activeServants.forEach((servant) => {
    const recordExists = db.financeRecords.some(
      (r) => r.campaignId === newCampaign.id && r.servantId === servant.id
    );
    if (!recordExists) {
      db.financeRecords.push({
        id: `fin-rec-${newCampaign.id}-${servant.id}`,
        campaignId: newCampaign.id,
        servantId: servant.id,
        servantCode: servant.code,
        servantName: servant.name,
        amount: 0,
        status: "NOT_PAID",
        notes: ""
      });
    }
  });

  // Create system notification
  db.notifications.push({
    id: `notif-${Date.now()}-finance-campaign`,
    title: `سجل مالي جديد: ${newCampaign.title}`,
    content: `تم فتح سجل مالي جديد لمتابعة الاشتراكات أو المساهمات بتاريخ ${newCampaign.date}.`,
    type: "SUCCESS",
    date: `${new Date().toISOString().split("T")[0]} 12:00`,
    isRead: false
  });

  saveDB(db).catch(console.error);
  auditLog("admin", "ADD", `إنشاء سجل مالي جديد: ${newCampaign.title}`);

  res.json({ success: true, campaign: newCampaign });
});


app.get("/api/honoring/records", async (req, res) => {
  const db = loadDB();
  db.honoringRecords = db.honoringRecords || [];
  res.json(db.honoringRecords);
});
app.get("/api/finance/records", async (req, res) => {
  const { campaignId } = req.query;
  const db = loadDB();
  db.financeRecords = db.financeRecords || [];

  if (!campaignId) {
    return res.json(db.financeRecords);
  }

  const servants = db.servants || [];

  // Sync: ensure every servant has a record for this campaign
  let changed = false;
  servants.forEach((servant) => {
    const recordIndex = db.financeRecords.findIndex(
      (r) => r.campaignId === campaignId && r.servantId === servant.id
    );

    if (recordIndex === -1) {
      db.financeRecords.push({
        id: `fin-rec-${campaignId}-${servant.id}`,
        campaignId: campaignId as string,
        servantId: servant.id,
        servantCode: servant.code,
        servantName: servant.name,
        amount: 0,
        status: "NOT_PAID",
        notes: ""
      });
      changed = true;
    } else {
      // In case servant name or code has changed, update it in the record
      const rec = db.financeRecords[recordIndex];
      if (rec.servantName !== servant.name || rec.servantCode !== servant.code) {
        rec.servantName = servant.name;
        rec.servantCode = servant.code;
        changed = true;
      }
    }
  });

  if (changed) {
    saveDB(db).catch(console.error);
  }

  // Filter records specifically for this campaign
  const campaignRecords = db.financeRecords.filter((r) => r.campaignId === campaignId);
  res.json(campaignRecords);
});

app.put("/api/finance/records/:id", async (req, res) => {
  const { id } = req.params;
  const { amount, status, notes } = req.body;
  const db = loadDB();

  db.financeRecords = db.financeRecords || [];
  const recordIndex = db.financeRecords.findIndex((r) => r.id === id);

  if (recordIndex === -1) {
    return res.status(404).json({ error: "السجل المالي غير موجود" });
  }

  const record = db.financeRecords[recordIndex];
  record.amount = Number(amount) || 0;
  record.status = status === "PAID" ? "PAID" : "NOT_PAID";
  record.notes = notes !== undefined ? notes.trim() : record.notes;
  record.paymentDate = record.status === "PAID" ? new Date().toISOString().split("T")[0] : undefined;

  saveDB(db).catch(console.error);
  auditLog(
    "admin", 
    "EDIT", 
    `تعديل مالية الخادم ${record.servantName}: المبلغ ${record.amount}، الحالة: ${record.status === "PAID" ? "واصل" : "غير واصل"}`
  );

  res.json({ success: true, record });
});

app.delete("/api/finance/campaigns/:id", async (req, res) => {
  const { id } = req.params;
  const db = loadDB();

  db.financeCampaigns = db.financeCampaigns || [];
  db.financeRecords = db.financeRecords || [];

  const campaignIndex = db.financeCampaigns.findIndex((c) => c.id === id);
  if (campaignIndex === -1) {
    return res.status(404).json({ error: "السجل المالي غير موجود" });
  }

  const campaignTitle = db.financeCampaigns[campaignIndex].title;

  // Remove campaign and its records
  db.financeCampaigns.splice(campaignIndex, 1);
  db.financeRecords = db.financeRecords.filter((r) => r.campaignId !== id);

  saveDB(db).catch(console.error);
  auditLog("admin", "DELETE", `حذف السجل المالي: ${campaignTitle}`);

  res.json({ success: true });
});

// -----------------------------------------------------------------
// LOCATIONS ENDPOINTS
// -----------------------------------------------------------------
app.get("/api/locations", async (req, res) => {
  const db = loadDB();
  res.json(db.locations);
});

app.post("/api/locations", async (req, res) => {
  const { name, requiredCount } = req.body;
  const db = loadDB();

  if (!name || requiredCount === undefined) {
    return res.status(400).json({ error: "الاسم والعدد المطلوب حقول إلزامية" });
  }

  const newLoc: Location = {
    id: `loc-${Date.now()}`,
    name,
    requiredCount: Number(requiredCount)
  };

  db.locations.push(newLoc);

  db.notifications.unshift({
    id: `notif-${Date.now()}`,
    title: "إضافة موقع خدمة جديد",
    content: `تم إنشاء موقع خدمة جديد (${name}) بالعدد المطلوب (${requiredCount} خادم).`,
    type: "SUCCESS",
    date: `${new Date().toISOString().split("T")[0]} 12:00`,
    isRead: false
  });

  saveDB(db).catch(console.error);
  auditLog("admin", "ADD", `إضافة موقع الخدمة: ${name} (العدد المطلوب: ${requiredCount})`);
  res.status(201).json(newLoc);
});

app.put("/api/locations/:id", async (req, res) => {
  const { id } = req.params;
  const { name, requiredCount } = req.body;
  const db = loadDB();

  const idx = db.locations.findIndex(l => l.id === id);
  if (idx === -1) return res.status(404).json({ error: "الموقع غير موجود" });

  db.locations[idx] = {
    ...db.locations[idx],
    name: name !== undefined ? name : db.locations[idx].name,
    requiredCount: requiredCount !== undefined ? Number(requiredCount) : db.locations[idx].requiredCount
  };

  db.notifications.unshift({
    id: `notif-${Date.now()}`,
    title: "تعديل موقع خدمة",
    content: `تم تحديث بيانات موقع الخدمة (${db.locations[idx].name}) بنجاح.`,
    type: "INFO",
    date: `${new Date().toISOString().split("T")[0]} 12:00`,
    isRead: false
  });

  saveDB(db).catch(console.error);
  auditLog("admin", "EDIT", `تعديل موقع الخدمة: ${db.locations[idx].name}`);
  res.json(db.locations[idx]);
});

app.delete("/api/locations/:id", async (req, res) => {
  const { id } = req.params;
  const db = loadDB();

  const idx = db.locations.findIndex(l => l.id === id);
  if (idx === -1) return res.status(404).json({ error: "الموقع غير موجود" });

  const loc = db.locations[idx];
  db.locations.splice(idx, 1);

  // Clean distributions to this location
  db.distributions = db.distributions.filter(d => d.locationId !== id);

  db.notifications.unshift({
    id: `notif-${Date.now()}`,
    title: "حذف موقع خدمة",
    content: `تم حذف موقع الخدمة (${loc.name}) وإلغاء توزيعه.`,
    type: "WARNING",
    date: `${new Date().toISOString().split("T")[0]} 12:00`,
    isRead: false
  });

  saveDB(db).catch(console.error);
  auditLog("admin", "DELETE", `حذف موقع الخدمة: ${loc.name}`);
  res.json({ success: true });
});

app.post("/api/locations/delete-all", async (req, res) => {
  const db = loadDB();
  db.locations = [];
  db.distributions = [];
  db.attendance = [];

  db.notifications.unshift({
    id: `notif-${Date.now()}`,
    title: "حذف جميع مواقع الميدان",
    content: "تم تفريغ وحذف سجل كافة مواقع الميدان وإلغاء جميع التوزيعات والحضور المرتبطة بها نهائياً.",
    type: "WARNING",
    date: `${new Date().toISOString().split("T")[0]} 12:00`,
    isRead: false
  });

  saveDB(db).catch(console.error);
  auditLog("admin", "DELETE", "حذف كافة مواقع ومحاور الميدان نهائياً");
  res.json({ success: true, message: "تم حذف جميع المواقع بنجاح" });
});

// -----------------------------------------------------------------
// SMART DISTRIBUTION ENGINE & MANUAL ASSIGNMENT
// -----------------------------------------------------------------
app.get("/api/distributions", async (req, res) => {
  try {
    const { dutyId } = req.query;
    const db = loadDB();
    const distributions = db.distributions || [];

    if (dutyId) {
      const list = distributions.filter((d: any) => d && d.dutyId === dutyId);
      return res.json(list);
    }
    res.json(distributions);
  } catch (error: any) {
    console.error("Error in GET /api/distributions:", error);
    res.status(500).json({ error: "Internal Server Error", details: error.message });
  }
});

// Smart distribution helper function to allocate servants equally across all locations
function distributeDuty(db: any, dutyId: string) {
  const duty = db.duties.find((d: any) => d.id === dutyId);
  if (!duty) return [];

  // 1. Get eligible active servants (ACTIVE, and not on leave or suspended or excluded)
  // And EXCLUDE those with "humanitarian" state = true from AUTO-distribution (they are assigned manually)
  const activeServants = db.servants.filter((s: any) => 
    s.status === ServantStatus.ACTIVE && 
    !s.humanitarian
  );

  // Servants with humanitarian needs will be kept unassigned or kept in their pinned/manual locations
  const humanitarianServants = db.servants.filter((s: any) => 
    s.status === ServantStatus.ACTIVE && 
    s.humanitarian
  );

  // 2. Clear previous non-pinned/non-manual distributions for this dutyId
  db.distributions = db.distributions.filter((d: any) => d.dutyId !== dutyId || d.isPinned);

  const currentDistributions = db.distributions.filter((d: any) => d.dutyId === dutyId);
  const distributedServantIds = new Set(currentDistributions.map((d: any) => d.servantId));

  // Eligible servants who are NOT yet assigned (not pinned)
  const pool = activeServants.filter((s: any) => !distributedServantIds.has(s.id));

  // 3. Get locations
  const locations = db.locations;
  if (locations.length === 0) return [];

  // Shuffle pool randomly first to ensure dynamic distribution on each run
  const shuffledPool = [...pool].sort(() => Math.random() - 0.5);

  // To distribute equally, we want each location to have as close to an equal number of servants as possible.
  // Let's keep track of current assignments per location (starting with pinned ones)
  const locationAssignments: Record<string, string[]> = {};
  locations.forEach((l: any) => {
    locationAssignments[l.id] = [];
  });

  currentDistributions.forEach((d: any) => {
    if (locationAssignments[d.locationId]) {
      locationAssignments[d.locationId].push(d.servantId);
    }
  });

  // Assign remaining pool to locations to make the total counts as equal as possible.
  // In each step, we assign the servant to the location with the minimum number of currently assigned servants.
  shuffledPool.forEach((servant: any) => {
    let minCount = Infinity;
    let minLocId = locations[0].id;

    // Shuffle locations order slightly to break ties randomly when counts are equal
    const shuffledLocations = [...locations].sort(() => Math.random() - 0.5);

    shuffledLocations.forEach((loc: any) => {
      const count = locationAssignments[loc.id].length;
      if (count < minCount) {
        minCount = count;
        minLocId = loc.id;
      }
    });

    locationAssignments[minLocId].push(servant.id);
  });

  // 4. Convert location assignments map back to DistributionItem database records
  const newDists: any[] = [];
  Object.entries(locationAssignments).forEach(([locId, servantIds]) => {
    servantIds.forEach((sId: any) => {
      // If already exists as pinned, skip re-adding it
      const exists = db.distributions.some((d: any) => d.dutyId === dutyId && d.servantId === sId);
      if (!exists) {
        newDists.push({
          id: `dist-${dutyId}-${sId}`,
          dutyId,
          servantId: sId,
          locationId: locId,
          isPinned: false
        });
      }
    });
  });

  // Automatically assign Humanitarian servants if they have a pinned location
  humanitarianServants.forEach((s: any) => {
    const exists = db.distributions.some((d: any) => d.dutyId === dutyId && d.servantId === s.id);
    if (!exists && s.pinnedLocationId) {
      newDists.push({
        id: `dist-${dutyId}-${s.id}`,
        dutyId,
        servantId: s.id,
        locationId: s.pinnedLocationId,
        isPinned: true
      });
    }
  });

  db.distributions.push(...newDists);

  // Generate blank attendance sheets for distributed servants to default to "UNMARKED" (isMarked: false)
  const allDutyDistributions = db.distributions.filter((d: any) => d.dutyId === dutyId);
  allDutyDistributions.forEach((dist: any) => {
    const attExists = db.attendance.some((a: any) => a.dutyId === dutyId && a.servantId === dist.servantId);
    if (!attExists) {
      db.attendance.push({
        id: `att-${dutyId}-${dist.servantId}`,
        dutyId,
        servantId: dist.servantId,
        locationId: dist.locationId,
        isPresent: false,
        isMarked: false, // Default is unmarked (neither present nor absent automatically)
        supervisorName: "التوزيع التلقائي",
        date: duty.date
      });
    }
  });

  return newDists;
}

// Smart distribution logic implementation
app.post("/api/distributions/auto", async (req, res) => {
  const { dutyId, allYear } = req.body;
  const db = loadDB();

  if (allYear) {
    if (db.locations.length === 0) {
      return res.status(400).json({ error: "يرجى إضافة مواقع الخدمة أولاً" });
    }

    let totalNewCount = 0;
    db.duties.forEach((duty: any) => {
      if (duty.isAutoDistributeDisabled) {
        console.log(`[Auto Distribute] Skipping duty ${duty.id} (${duty.label}) as auto-distribute is disabled.`);
        return;
      }
      const newDists = distributeDuty(db, duty.id);
      totalNewCount += newDists.length;
    });

    db.notifications.unshift({
      id: `notif-${Date.now()}-dist-year`,
      title: "اكتمال التوزيع التلقائي السنوي",
      content: `قام النظام بتوزيع الخدام بالتساوي على كافة مواقع الخدمة لجميع واجبات السنة كاملة بنجاح (مع تخطي الواجبات المستثناة).`,
      type: "SUCCESS",
      date: `${new Date().toISOString().split("T")[0]} 12:00`,
      isRead: false
    });

    saveDB(db).catch(console.error);
    auditLog("admin", "DISTRIBUTE", "تنفيذ محرك التوزيع التلقائي السنوي بالتساوي لكافة واجبات السنة");

    return res.json({ success: true, message: "تم التوزيع التلقائي لكافة واجبات السنة بنجاح", count: totalNewCount });
  }

  if (!dutyId) return res.status(400).json({ error: "واجب الخدمة مطلوب لتنفيذ التوزيع الذكي" });

  const duty = db.duties.find(d => d.id === dutyId);
  if (!duty) return res.status(404).json({ error: "الواجب غير موجود" });

  if (duty.isAutoDistributeDisabled) {
    return res.status(400).json({ error: "تم إلغاء ميزة التوزيع التلقائي لهذا الواجب خصيصاً" });
  }

  if (db.locations.length === 0) {
    return res.status(400).json({ error: "يرجى إضافة مواقع الخدمة أولاً" });
  }

  const newDists = distributeDuty(db, dutyId);

  db.notifications.unshift({
    id: `notif-${Date.now()}-dist`,
    title: `اكتمال التوزيع الذكي لواجب ${duty.label}`,
    content: `قام النظام بتوزيع الخدام بالتساوي وتعميم الكشوف بنجاح.`,
    type: "SUCCESS",
    date: `${new Date().toISOString().split("T")[0]} 12:00`,
    isRead: false
  });

  saveDB(db).catch(console.error);
  auditLog("admin", "DISTRIBUTE", `تنفيذ محرك التوزيع بالتساوي للخدام لـ: ${duty.label}`);

  res.json({ success: true, message: "تم التوزيع بالتساوي للخدام بنجاح", count: newDists.length });
});

// Custom equal numerical distribution by target count and date
app.post("/api/distributions/auto-by-count", async (req, res) => {
  const { date, count } = req.body;
  const db = loadDB();

  if (!date) {
    return res.status(400).json({ error: "تاريخ الواجب مطلوب لإجراء التوزيع بالتساوي" });
  }

  const targetCount = Number(count);
  if (isNaN(targetCount) || targetCount <= 0) {
    return res.status(400).json({ error: "عدد الخدام يجب أن يكون رقماً صحيحاً أكبر من الصفر" });
  }

  // 1. Find or create the duty for this date
  const dateStr = date.trim();
  let duty = db.duties.find((d: any) => d.date === dateStr);
  let isNewDuty = false;

  if (!duty) {
    const dObj = new Date(dateStr);
    const daysMap: Record<number, string> = {
      0: "Sunday",
      1: "Monday",
      2: "Tuesday",
      3: "Wednesday",
      4: "Thursday",
      5: "Friday",
      6: "Saturday"
    };
    const dayOfWeek = daysMap[dObj.getDay()] || "Special";
    duty = {
      id: `duty-special-${Date.now()}`,
      date: dateStr,
      dayOfWeek: dayOfWeek,
      year: dObj.getFullYear() || 2026,
      label: `واجب توزيع عددي (${targetCount} خادم)`
    };
    db.duties.push(duty);
    isNewDuty = true;
  }

  const dutyId = duty.id;

  if (db.locations.length === 0) {
    return res.status(400).json({ error: "يرجى إضافة مواقع الخدمة أولاً قبل التوزيع" });
  }

  // 2. Select eligible active servants
  const activeServants = db.servants.filter((s: any) => 
    s.status === "ACTIVE" && 
    !s.humanitarian
  );

  // Clear previous non-pinned distributions for this dutyId to get correct available count
  db.distributions = db.distributions.filter((d: any) => d.dutyId !== dutyId || d.isPinned);

  const currentDistributions = db.distributions.filter((d: any) => d.dutyId === dutyId);
  const distributedServantIds = new Set(currentDistributions.map((d: any) => d.servantId));

  // Eligible pool of servants
  const pool = activeServants.filter((s: any) => !distributedServantIds.has(s.id));

  // Shuffle pool and take exactly targetCount (or all available if targetCount is larger)
  const selectedPool = [...pool].sort(() => Math.random() - 0.5).slice(0, targetCount);

  if (selectedPool.length === 0) {
    return res.status(400).json({ error: "لا يوجد خدام نشطين متاحين للتوزيع" });
  }

  // 3. Initialize location assignments starting with pinned ones
  const locations = db.locations;
  const locationAssignments: Record<string, string[]> = {};
  locations.forEach((l: any) => {
    locationAssignments[l.id] = [];
  });

  currentDistributions.forEach((d: any) => {
    if (locationAssignments[d.locationId]) {
      locationAssignments[d.locationId].push(d.servantId);
    }
  });

  // Distribute selectedPool to locations equally
  selectedPool.forEach((servant: any) => {
    let minCount = Infinity;
    let minLocId = locations[0].id;

    // Shuffle locations to break ties randomly
    const shuffledLocations = [...locations].sort(() => Math.random() - 0.5);

    shuffledLocations.forEach((loc: any) => {
      const currentCount = locationAssignments[loc.id].length;
      if (currentCount < minCount) {
        minCount = currentCount;
        minLocId = loc.id;
      }
    });

    locationAssignments[minLocId].push(servant.id);
  });

  // Convert assignments back to database records
  const newDists: any[] = [];
  Object.entries(locationAssignments).forEach(([locId, servantIds]) => {
    servantIds.forEach((sId: any) => {
      const exists = db.distributions.some((d: any) => d.dutyId === dutyId && d.servantId === sId);
      if (!exists) {
        newDists.push({
          id: `dist-${dutyId}-${sId}`,
          dutyId,
          servantId: sId,
          locationId: locId,
          isPinned: false
        });
      }
    });
  });

  db.distributions.push(...newDists);

  // Generate blank attendance sheets
  const allDutyDistributions = db.distributions.filter((d: any) => d.dutyId === dutyId);
  allDutyDistributions.forEach((dist: any) => {
    const attExists = db.attendance.some((a: any) => a.dutyId === dutyId && a.servantId === dist.servantId);
    if (!attExists) {
      db.attendance.push({
        id: `att-${dutyId}-${dist.servantId}`,
        dutyId,
        servantId: dist.servantId,
        locationId: dist.locationId,
        isPresent: false,
        isMarked: false,
        supervisorName: "التوزيع التلقائي العددي",
        date: duty.date
      });
    }
  });

  db.notifications.unshift({
    id: `notif-${Date.now()}-dist-count`,
    title: `اكتمال التوزيع العددي المساوي لواجب ${duty.label}`,
    content: `تم توزيع ${selectedPool.length} خادم بالتساوي على ${locations.length} موقع لمناسبة تاريخ ${date}.`,
    type: "SUCCESS",
    date: `${new Date().toISOString().split("T")[0]} 12:00`,
    isRead: false
  });

  saveDB(db).catch(console.error);
  auditLog("admin", "DISTRIBUTE", `تنفيذ التوزيع العددي المتساوي لـ ${selectedPool.length} خادم في تاريخ ${date}`);

  

  res.json({
    success: true,
    message: "تم التوزيع العددي المتساوي للخدام بنجاح",
    dutyId,
    count: selectedPool.length,
    totalAvailable: pool.length
  });
});

// Random field distribution endpoint
app.post("/api/distributions/random-field", async (req, res) => {
  const { dutyId, count } = req.body;
  const db = loadDB();

  if (!dutyId) {
    return res.status(400).json({ error: "واجب الخدمة مطلوب" });
  }

  const duty = db.duties.find(d => d.id === dutyId);
  if (!duty) {
    return res.status(404).json({ error: "واجب الخدمة غير موجود" });
  }

  const reqCount = Number(count) || 20;

  // Find all attendance records for this duty where servant is present
  const presentAttendance = db.attendance.filter(a => a.dutyId === dutyId && a.isPresent);
  const presentServantIds = presentAttendance.map(a => a.servantId);

  // Filter actual servants in the DB
  const presentServants = db.servants.filter(s => presentServantIds.includes(s.id));

  if (presentServants.length === 0) {
    return res.status(400).json({ error: "لا يوجد خُدّام حاضرون في الميدان لهذا الواجب حالياً لتوزيعهم عشوائياً. يرجى تسجيل الحضور أولاً." });
  }

  // Shuffle and pick count
  const shuffled = [...presentServants].sort(() => 0.5 - Math.random());
  const selectedCount = Math.min(shuffled.length, reqCount);
  const selectedServants = shuffled.slice(0, selectedCount);

  // Clear previous sector leader attributes for all servants
  db.servants.forEach(s => {
    s.isSectorLeader = false;
    s.sectorRole = undefined;
  });

  // Assign one of the selected as "مسؤول قاطع"
  let leaderName = "";
  if (selectedServants.length > 0) {
    const leaderIndex = Math.floor(Math.random() * selectedServants.length);
    const leaderServant = db.servants.find(s => s.id === selectedServants[leaderIndex].id);
    if (leaderServant) {
      leaderServant.isSectorLeader = true;
      leaderServant.sectorRole = "مسؤول قاطع";
      leaderName = leaderServant.name;
      // also update local references in our response array
      selectedServants.forEach(s => {
        if (s.id === leaderServant.id) {
          s.isSectorLeader = true;
          s.sectorRole = "مسؤول قاطع";
        }
      });
    }
  }

  db.notifications.unshift({
    id: `notif-${Date.now()}-rand-dist`,
    title: `التوزيع العشوائي لمسؤول القاطع: واجب ${duty.label}`,
    content: `تم سحب عينة عشوائية مكونة من ${selectedCount} خادم من الحاضرين، وتعيين الخادم (${leaderName}) كمسؤول للقاطع.`,
    type: "SUCCESS",
    date: `${new Date().toISOString().split("T")[0]} 12:00`,
    isRead: false
  });

  saveDB(db).catch(console.error);
  auditLog("admin", "DISTRIBUTE", `سحب عينة عشوائية من الحاضرين لـ ${duty.label} وتعيين ${leaderName} كمسؤول قاطع`);

  res.json({
    success: true,
    selectedCount,
    leaderName,
    selectedServants
  });
});

// Update/pin custom assignments immediately (Drag & Drop or Manual Selection)
app.post("/api/distributions/update", async (req, res) => {
  const { dutyId, servantId, locationId, isPinned } = req.body;
  if (!dutyId || !servantId || !locationId) {
    return res.status(400).json({ error: "معرّف الواجب، الخادم والموقع حقول إلزامية" });
  }

  const db = loadDB();
  const duty = db.duties.find(d => d.id === dutyId);
  if (!duty) return res.status(404).json({ error: "الواجب غير موجود" });

  const idx = db.distributions.findIndex(d => d.dutyId === dutyId && d.servantId === servantId);
  const servant = db.servants.find(s => s.id === servantId);
  const location = db.locations.find(l => l.id === locationId);

  const locName = location ? location.name : "غير محدد";

  if (idx !== -1) {
    db.distributions[idx].locationId = locationId;
    db.distributions[idx].isPinned = isPinned !== undefined ? isPinned : true;
  } else {
    db.distributions.push({
      id: `dist-${dutyId}-${servantId}`,
      dutyId,
      servantId,
      locationId,
      isPinned: isPinned !== undefined ? isPinned : true
    });
  }

  // Sync attendance location as well
  const attIdx = db.attendance.findIndex(a => a.dutyId === dutyId && a.servantId === servantId);
  if (attIdx !== -1) {
    db.attendance[attIdx].locationId = locationId;
  } else {
    db.attendance.push({
      id: `att-${dutyId}-${servantId}`,
      dutyId,
      servantId,
      locationId,
      isPresent: false,
      isMarked: false, // Default to unmarked (neither present nor absent automatically)
      supervisorName: "تعديل يدوي",
      date: duty.date
    });
  }

  saveDB(db).catch(console.error);
  auditLog("supervisor", "EDIT", `نقل الخادم ${servant ? servant.name : servantId} يدوياً إلى موقع: ${locName}`);

  res.json({ success: true, message: "تم تعديل التوزيع وحفظه فوراً" });
});

app.post("/api/distributions/reset", async (req, res) => {
  const { dutyId } = req.body;
  if (!dutyId) return res.status(400).json({ error: "الواجب مطلوب لإعادة الضبط" });

  const db = loadDB();
  db.distributions = db.distributions.filter(d => d.dutyId !== dutyId);
  db.attendance = db.attendance.filter(a => a.dutyId !== dutyId);
  saveDB(db).catch(console.error);

  auditLog("supervisor", "DELETE", `إعادة ضبط وتصفير كشوف توزيع الواجب: ${dutyId}`);

  res.json({ success: true, message: "تم إعادة ضبط توزيع الواجب بنجاح" });
});

// -----------------------------------------------------------------
// ATTENDANCE ENDPOINTS
// -----------------------------------------------------------------
app.get("/api/attendance", async (req, res) => {
  const { dutyId } = req.query;
  const db = loadDB();

  if (dutyId) {
    const list = db.attendance.filter(a => a.dutyId === dutyId);
    return res.json(list);
  }
  res.json(db.attendance);
});

// Record or toggle attendance
app.delete("/api/attendance/duty/:dutyId/servant/:servantId", async (req, res) => {
  const { dutyId, servantId } = req.params;
  const db = loadDB();
  db.attendance = db.attendance.filter(a => !(a.dutyId === dutyId && a.servantId === servantId));
  saveDB(db).catch(console.error);
  setHasPendingChanges(true); // From auto-save
  res.json({ success: true });
});

app.delete("/api/attendance/duty/:dutyId", async (req, res) => {
  const { dutyId } = req.params;
  const db = loadDB();
  const initialCount = db.attendance.length;
  db.attendance = db.attendance.filter(a => a.dutyId !== dutyId);
  const deletedCount = initialCount - db.attendance.length;

  // Log
  const targetDuty = db.duties.find(d => d.id === dutyId);
  const dutyDate = targetDuty ? targetDuty.date : dutyId;
  console.log("[DELETE_ALL_ATTENDANCE]");
  console.log(`Date=${dutyDate}`);
  console.log(`DeletedRecords=${deletedCount}`);

  saveDB(db).catch(console.error);
  setHasPendingChanges(true); // From auto-save

  // Note: we don't push automatically here to avoid spamming sheets, it will happen automatically.
  // The user prompt says: "نفّذ saveDB() مرة واحدة فقط"
  
  auditLog("admin", "DELETE", `حذف جميع سجلات الحضور ليوم: ${dutyDate}`);
  res.json({ success: true, deletedCount });
});

app.post("/api/attendance", async (req, res) => {
  const { dutyId, servantId, locationId, isPresent, supervisorName, accessToken } = req.body;
  if (!dutyId || !servantId) return res.status(400).json({ error: "الواجب والخادم مطلوبان" });

  const db = loadDB();
  const duty = db.duties.find(d => d.id === dutyId);
  if (!duty) return res.status(404).json({ error: "الواجب غير موجود" });

  const servant = db.servants.find(s => s.id === servantId);
  if (!servant) return res.status(404).json({ error: "الخادم غير موجود" });

  const idx = db.attendance.findIndex(a => a.dutyId === dutyId && a.servantId === servantId);
  const timeStr = isPresent ? new Date().toTimeString().split(" ")[0].substring(0, 5) : undefined;

  let finalLocationId = locationId;
  if (isPresent && (!finalLocationId || finalLocationId === "unassigned")) {
    if (servant.pinnedLocationId && db.locations.some(l => l.id === servant.pinnedLocationId)) {
      finalLocationId = servant.pinnedLocationId;
    } else if (db.locations && db.locations.length > 0) {
      // Smart: count how many present servants are in each location for this duty to balance field loads
      const counts: Record<string, number> = {};
      db.locations.forEach(l => {
        counts[l.id] = 0;
      });
      db.attendance.forEach((a: any) => {
        if (a.dutyId === dutyId && a.isPresent && counts[a.locationId] !== undefined) {
          counts[a.locationId]++;
        }
      });
      // Pick the least-populated location
      let minLocId = db.locations[0].id;
      let minCount = counts[minLocId];
      for (let i = 1; i < db.locations.length; i++) {
        const locId = db.locations[i].id;
        if (counts[locId] < minCount) {
          minCount = counts[locId];
          minLocId = locId;
        }
      }
      finalLocationId = minLocId;
    }
  }

  if (!finalLocationId) {
    finalLocationId = "unassigned";
  }

  if (idx !== -1) {
    db.attendance[idx].isPresent = !!isPresent;
    db.attendance[idx].time = timeStr;
    db.attendance[idx].locationId = finalLocationId || db.attendance[idx].locationId;
    db.attendance[idx].supervisorName = supervisorName || "الشيخ خالد (المشرف)";
    db.attendance[idx].isMarked = true; // Mark as manually recorded
  } else {
    db.attendance.push({
      id: `att-${dutyId}-${servantId}`,
      dutyId,
      servantId,
      locationId: finalLocationId,
      isPresent: !!isPresent,
      time: timeStr,
      supervisorName: supervisorName || "الشيخ خالد (المشرف)",
      date: duty.date,
      isMarked: true // Mark as manually recorded
    });
  }

  // Trigger automated scan of consecutive absences immediately on attendance updates!
  // This satisfies the Real-time alerts requirement
  saveDB(db).catch(console.error);
  generateAndPushReportsToGoogleSheet(accessToken).catch(console.error);
  const { dbChanged } = runSystemScans(db, supervisorName || "نظام الحضور");

  auditLog(
    supervisorName || "supervisor", 
    "EDIT", 
    `تسجيل الخادم ${servant.name} كـ (${isPresent ? "حاضر" : "غائب"}) في واجب ${duty.label}`
  );

  res.json({ success: true, isPresent, time: timeStr, dbChanged });
});

app.post("/api/attendance/bulk-absent", async (req, res) => {
  const { dutyId, unmarkedServants, supervisorName, accessToken } = req.body;
  if (!dutyId || !unmarkedServants || !Array.isArray(unmarkedServants)) {
    return res.status(400).json({ error: "الواجب وقائمة الخدم مطلوبان" });
  }

  const db = loadDB();
  const duty = db.duties.find(d => d.id === dutyId);
  if (!duty) return res.status(404).json({ error: "الواجب غير موجود" });

  const todayBaghdad = getBaghdadDate();
  const todayStr = todayBaghdad.toISOString().split("T")[0];

  // Try to fetch latest leaves from "سجل الاجازات 2026" sheet to be completely up to date
  let sheetLeaves: any[] = [];
  try {
    const email = db.sheetsSyncConfig?.clientEmail;
    const pkey = db.sheetsSyncConfig?.privateKey;
    const rawId = resolveSpreadsheetId();
    
    if (rawId && ((email && pkey) || accessToken)) {
      if (!rawId) throw new Error("لم يتم إعداد معرف جدول البيانات Google Sheets. يرجى إعداده من الإعدادات قبل المزامنة.");
  const cleanId = rawId.trim();
      const match = cleanId.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
      const spreadsheetId = match && match[1] ? match[1] : cleanId;
  console.log("[SYNC_UNIFIED_SPREADSHEET_ID]", spreadsheetId);

      let authClient: any = null;
      if (email && pkey) {
        authClient = new google.auth.JWT({
          email: email,
          key: pkey.replace(/\\n/g, '\n'),
          scopes: ['https://www.googleapis.com/auth/spreadsheets']
        });
        await authClient.authorize();
      } else if (accessToken) {
        const oauth2Client = new google.auth.OAuth2();
        oauth2Client.setCredentials({ access_token: accessToken });
        authClient = oauth2Client;
      }

      if (authClient) {
        console.log("[TRACE_PUSH_STARTED]");
  console.log("[TRACE_SPREADSHEET_ID]", spreadsheetId);
    console.log("[TRACE_ACCESS_TOKEN]", !!accessToken);
  // customAccessToken might be the name here
  console.log("[TRACE_AUTH_CLIENT]", !!authClient);
  if (authClient) console.log("[TRACE_GOOGLE_CONNECTED]");
  const sheets = google.sheets({ version: "v4", auth: authClient });
  const meta = await withSheetsRetry(() => sheets.spreadsheets.get({ spreadsheetId }));
        const hasLeavesTab = meta.data.sheets?.some(s => s.properties?.title === "سجل الاجازات 2026");
        
        if (hasLeavesTab) {
          const lResponse = await withSheetsRetry(() => sheets.spreadsheets.values.get({
            spreadsheetId,
            range: "'سجل الاجازات 2026'!A1:F1000"
          }));
          const leavesRows = lResponse.data.values || [];
          if (leavesRows.length > 1) {
            // Clear and sync local sheets leaves
            db.leaves = db.leaves.filter(l => !l.id.startsWith("sheet-"));
            
            for (let i = 1; i < leavesRows.length; i++) {
              const row = leavesRows[i];
              if (!row || row.length < 3) continue;
            const code = String(row[0] || "").trim();
            const name = String(row[1] || "").trim();
            const startDateStr = String(row[2] || "").trim();
            const durationStr = String(row[3] || "").trim();
            const endDateStr = String(row[4] || "").trim();
            const details = String(row[5] || "").trim();
            
            if (!code) continue;
            const servant = db.servants.find(s => s.code === code);
            if (!servant) continue;

            const startD = parseSheetDate(startDateStr);
            let endD = parseSheetDate(endDateStr);
            let durationDays = 1;

            if (startD) {
              if (endD) {
                durationDays = Math.max(1, Math.round((endD.getTime() - startD.getTime()) / (1000 * 60 * 60 * 24)));
              } else {
                const parsedDuration = parseInt(durationStr);
                const days = isNaN(parsedDuration) || parsedDuration <= 0 ? 1 : parsedDuration;
                durationDays = days;
                endD = new Date(startD.getTime() + days * 24 * 60 * 60 * 1000);
              }

              const formattedStart = startD.toISOString().split("T")[0];
              const formattedEnd = endD.toISOString().split("T")[0];
              
              let status: "ACTIVE" | "COMPLETED" | "PENDING" = "PENDING";
              if (todayBaghdad.getTime() >= startD.getTime() && todayBaghdad.getTime() <= endD.getTime()) {
                status = "ACTIVE";
              } else if (todayBaghdad.getTime() > endD.getTime()) {
                status = "COMPLETED";
              }
              
              const syncedLeave: Leave = {
                id: `sheet-${code}-${formattedStart}`,
                servantId: servant.id,
                servantCode: servant.code,
                servantName: servant.name,
                startDate: formattedStart,
                endDate: formattedEnd,
                durationDays: durationDays,
                reason: details || "إجازة رسمية عبر Google Sheets",
                status: status
              };
              db.leaves.push(syncedLeave);
              sheetLeaves.push(syncedLeave);
            }
          }
        }
      }
    }
  }
  } catch (sheetErr: any) {
    console.error("[Bulk Mark Absent] Could not read latest leaves from Sheets:", sheetErr.message || sheetErr);
    // Fall back to local db.leaves
    sheetLeaves = db.leaves || [];
  }

  const timestampStr = new Date().toISOString();
  const currentTimeStr = new Date().toLocaleTimeString("ar-SA", { hour: "2-digit", minute: "2-digit" });

  let markedCount = 0;
  let skippedCount = 0;

  unmarkedServants.forEach(({ servantId, locationId }) => {
    const servant = db.servants.find(s => s.id === servantId);
    if (!servant) return;

    // Check if on official leave today (Baghdad time)
    const isOnLeaveToday = sheetLeaves.some(l => {
      if (String(l.servantId) !== String(servant.id)) return false;
      const startD = parseSheetDate(l.startDate);
      const endD = parseSheetDate(l.endDate);
      if (startD && endD) {
        return todayBaghdad.getTime() >= startD.getTime() && todayBaghdad.getTime() <= endD.getTime();
      }
      return false;
    });

    if (isOnLeaveToday) {
      skippedCount++;
      return;
    }

    const idx = db.attendance.findIndex(a => a.dutyId === dutyId && a.servantId === servantId);
    if (idx !== -1) {
      db.attendance[idx] = {
        ...db.attendance[idx],
        isPresent: false,
        supervisorName: supervisorName || "النظام التلقائي",
        isMarked: true,
        date: db.attendance[idx].date || duty.date,
        time: db.attendance[idx].time || currentTimeStr,
        timestamp: timestampStr
      };
    } else {
      db.attendance.push({
        id: `att_${Date.now()}_${Math.random().toString(36).substr(2, 5)}`,
        dutyId,
        servantId,
        locationId: locationId || "unassigned",
        isPresent: false,
        supervisorName: supervisorName || "النظام التلقائي",
        isMarked: true,
        date: duty.date,
        time: currentTimeStr,
        timestamp: timestampStr
      });
    }
    markedCount++;
  });

  console.log("[MARK_REMAINING_ABSENT]");
  console.log("Date=" + duty.date);
  console.log("Count=" + markedCount);

  saveDB(db).catch(console.error);

  // Push to permanent log directly as requested
  generateAndPushReportsToGoogleSheet(accessToken).catch(console.error);
  const { dbChanged } = runSystemScans(db, supervisorName || "نظام الحضور");

  

  auditLog(
    supervisorName || "system",
    "EDIT",
    `تسجيل غياب جماعي للواجب ${duty.label}: تم قيد ${markedCount} غياب، وتخطي ${skippedCount} مجازين`
  );

  res.json({
    success: true,
    message: `اكتمل تسجيل الغياب التلقائي لليوم: تم تسجيل ${markedCount} غائب، وتخطي ${skippedCount} مجاز رسمي تزامناً مع فحص شيت سجل الاجازات 2026.`,
    db,
    dbChanged
  });
});

// -----------------------------------------------------------------
// LEAVES ENDPOINTS
// -----------------------------------------------------------------
app.get("/api/leaves", async (req, res) => {
  const db = loadDB();
  res.json(db.leaves);
});

app.post("/api/leaves", async (req, res) => {
  const { servantId, reason, startDate, durationDays } = req.body;
  const db = loadDB();

  const servant = db.servants.find(s => s.id === servantId);
  if (!servant) return res.status(404).json({ error: "الخادم غير موجود" });

  const start = new Date(startDate);
  const end = new Date(start);
  end.setDate(start.getDate() + Number(durationDays));
  const endDateStr = end.toISOString().split("T")[0];

  const newLeave: Leave = {
    id: `leave-${Date.now()}`,
    servantId,
    servantCode: servant.code,
    servantName: servant.name,
    reason,
    startDate,
    durationDays: Number(durationDays),
    endDate: endDateStr,
    status: "ACTIVE"
  };

  db.leaves.push(newLeave);

  // Update servant status to "ON_LEAVE"
  servant.status = ServantStatus.ON_LEAVE;
  servant.statusDate = startDate || new Date().toISOString().split("T")[0];

  // Add notification
  db.notifications.unshift({
    id: `notif-${Date.now()}-leav`,
    title: "منح إجازة خادم",
    content: `تم منح الخادم ${servant.name} (كود: ${servant.code}) إجازة لمدة ${durationDays} يوم تبدأ من ${startDate} وتنتهي في ${endDateStr}.`,
    type: "LEAVE",
    date: `${new Date().toISOString().split("T")[0]} 12:00`,
    isRead: false
  });

  saveDB(db).catch(console.error);
  auditLog("admin", "ADD", `منح إجازة للخادم: ${servant.name} (السبب: ${reason})`);
  res.status(201).json(newLeave);
});

app.post("/api/leaves/request", async (req, res) => {
  const { servantId, reason, startDate, durationDays } = req.body;
  const db = loadDB();

  const servant = db.servants.find(s => s.id === servantId);
  if (!servant) return res.status(404).json({ error: "الخادم غير موجود" });

  const start = new Date(startDate);
  const end = new Date(start);
  end.setDate(start.getDate() + Number(durationDays));
  const endDateStr = end.toISOString().split("T")[0];

  const newLeave: Leave = {
    id: `leave-${Date.now()}`,
    servantId,
    servantCode: servant.code,
    servantName: servant.name,
    reason,
    startDate,
    durationDays: Number(durationDays),
    endDate: endDateStr,
    status: "PENDING"
  };

  db.leaves.push(newLeave);

  // Add system notification for supervisors
  db.notifications.unshift({
    id: `notif-${Date.now()}-leav-req`,
    title: "طلب إجازة جديد معلّق",
    content: `قام الخادم ${servant.name} (كود: ${servant.code}) بتقديم طلب إجازة لمدة ${durationDays} يوم تبدأ من ${startDate} وبانتظار الموافقة.`,
    type: "LEAVE",
    date: `${new Date().toISOString().split("T")[0]} 12:00`,
    isRead: false
  });

  saveDB(db).catch(console.error);
  auditLog(servant.name, "ADD", `تقديم طلب إجازة معلّق: ${reason}`);
  res.status(201).json(newLeave);
});

app.post("/api/leaves/:id/approve", async (req, res) => {
  const { id } = req.params;
  const { supervisorName } = req.body;
  const db = loadDB();

  const leave = db.leaves.find(l => l.id === id);
  if (!leave) return res.status(404).json({ error: "طلب الإجازة غير موجود" });

  if (leave.status !== "PENDING") {
    return res.status(400).json({ error: "طلب الإجازة ليس معلقاً" });
  }

  const servant = db.servants.find(s => s.id === leave.servantId);
  if (!servant) return res.status(404).json({ error: "الخادم غير موجود" });

  // Update status of leave to ACTIVE
  leave.status = "ACTIVE";

  // Update servant status to ON_LEAVE
  servant.status = ServantStatus.ON_LEAVE;
  servant.statusDate = leave.startDate || new Date().toISOString().split("T")[0];

  // Notification
  db.notifications.unshift({
    id: `notif-${Date.now()}-leav-app`,
    title: "الموافقة على طلب إجازة",
    content: `تمت الموافقة على طلب إجازة الخادم ${servant.name} (كود: ${servant.code}) لمدة ${leave.durationDays} يوم تبدأ من ${leave.startDate}.`,
    type: "SUCCESS",
    date: `${new Date().toISOString().split("T")[0]} 12:00`,
    isRead: false
  });

  saveDB(db).catch(console.error);
  auditLog(supervisorName || "admin", "EDIT", `الموافقة على إجازة الخادم: ${servant.name}`);
  res.json({ success: true, leave });
});

app.post("/api/leaves/:id/reject", async (req, res) => {
  const { id } = req.params;
  const { supervisorName } = req.body;
  const db = loadDB();

  const leave = db.leaves.find(l => l.id === id);
  if (!leave) return res.status(404).json({ error: "طلب الإجازة غير موجود" });

  if (leave.status !== "PENDING") {
    return res.status(400).json({ error: "طلب الإجازة ليس معلقاً" });
  }

  const servant = db.servants.find(s => s.id === leave.servantId);
  if (!servant) return res.status(404).json({ error: "الخادم غير موجود" });

  // Update status of leave to REJECTED
  leave.status = "REJECTED";

  // Notification
  db.notifications.unshift({
    id: `notif-${Date.now()}-leav-rej`,
    title: "رفض طلب إجازة",
    content: `تم رفض طلب إجازة الخادم ${servant.name} (كود: ${servant.code}) المقدمة من تاريخ ${leave.startDate}.`,
    type: "WARNING",
    date: `${new Date().toISOString().split("T")[0]} 12:00`,
    isRead: false
  });

  saveDB(db).catch(console.error);
  auditLog(supervisorName || "admin", "EDIT", `رفض إجازة الخادم: ${servant.name}`);
  res.json({ success: true, leave });
});

// -----------------------------------------------------------------
// SUSPENSIONS ENDPOINTS
// -----------------------------------------------------------------
app.get("/api/suspensions", async (req, res) => {
  const db = loadDB();
  res.json(db.suspensions);
});

app.post("/api/suspensions", async (req, res) => {
  const { servantId, reason, startDate, durationMonths } = req.body;
  const db = loadDB();

  const servant = db.servants.find(s => s.id === servantId);
  if (!servant) return res.status(404).json({ error: "الخادم غير موجود" });

  const start = new Date(startDate);
  const end = new Date(start);
  end.setMonth(start.getMonth() + Number(durationMonths));
  const endDateStr = end.toISOString().split("T")[0];

  const newSusp: Suspension = {
    id: `susp-${Date.now()}`,
    servantId,
    servantCode: servant.code,
    servantName: servant.name,
    reason,
    startDate,
    durationMonths: Number(durationMonths),
    endDate: endDateStr,
    status: "ACTIVE"
  };

  db.suspensions.push(newSusp);

  // Update servant status
  servant.status = ServantStatus.SUSPENDED;
  servant.statusDate = startDate || new Date().toISOString().split("T")[0];

  db.notifications.unshift({
    id: `notif-${Date.now()}-susp`,
    title: "إصدار قرار توقيف خادم",
    content: `تم إيقاف الخادم ${servant.name} (كود: ${servant.code}) لمدة ${durationMonths} أشهر تبدأ من ${startDate} وتنتهي في ${endDateStr}.`,
    type: "SUSPENSION",
    date: `${new Date().toISOString().split("T")[0]} 12:00`,
    isRead: false
  });

  saveDB(db).catch(console.error);
  auditLog("admin", "ADD", `إصدار توقيف للخادم: ${servant.name} (السبب: ${reason})`);
  res.status(201).json(newSusp);
});

// Approve Reactivation of suspended servant after suspension duration ends
const handleSuspensionApprove = (req: any, res: any) => {
  const { id } = req.params;
  const db = loadDB();

  const suspIdx = db.suspensions.findIndex(s => s.id === id);
  if (suspIdx === -1) return res.status(404).json({ error: "سجل التوقيف غير موجود" });

  const susp = db.suspensions[suspIdx];
  susp.status = "COMPLETED";

  const servant = db.servants.find(s => s.id === susp.servantId);
  if (servant) {
    servant.status = ServantStatus.ACTIVE;
    servant.statusDate = new Date().toISOString().split("T")[0];
  }

  db.notifications.unshift({
    id: `notif-${Date.now()}-susp-react`,
    title: "موافقة على تفعيل خادم موقوف",
    content: `تمت الموافقة من المشرف على تفعيل وإعادة ملف الخادم ${susp.servantName} (كود: ${susp.servantCode}) إلى الخدمة الفعالة.`,
    type: "SUCCESS",
    date: `${new Date().toISOString().split("T")[0]} 12:00`,
    isRead: false
  });

  saveDB(db).catch(console.error);
  auditLog("supervisor", "EDIT", `الموافقة على إنهاء توقيف وتفعيل الخادم: ${susp.servantName}`);
  res.json({ success: true, message: "تم تفعيل الخادم بنجاح وموافقة المشرف" });
};

app.put("/api/suspensions/:id/approve", handleSuspensionApprove);
app.post("/api/suspensions/:id/approve", handleSuspensionApprove);

// -----------------------------------------------------------------
// EXCLUDED ENDPOINTS
// -----------------------------------------------------------------
app.get("/api/excluded", async (req, res) => {
  const db = loadDB();
  res.json(db.excluded);
});

app.post("/api/excluded", async (req, res) => {
  const { servantId, reason, date, notes } = req.body;
  const db = loadDB();

  const servant = db.servants.find(s => s.id === servantId);
  if (!servant) return res.status(404).json({ error: "الخادم غير موجود" });

  const newExcluded: ExcludedServant = {
    id: `excl-${Date.now()}`,
    servantId,
    servantCode: servant.code,
    servantName: servant.name,
    reason,
    date: date || new Date().toISOString().split("T")[0],
    notes
  };

  db.excluded.push(newExcluded);

  // Set servant status to EXCLUDED
  servant.status = ServantStatus.EXCLUDED;
  servant.statusDate = date || new Date().toISOString().split("T")[0];

  db.notifications.unshift({
    id: `notif-${Date.now()}-excl`,
    title: "استبعاد خادم من الهيئة",
    content: `تم نقل الخادم ${servant.name} (كود: ${servant.code}) لسجل المستبعدين بصفة نهائية.`,
    type: "WARNING",
    date: `${new Date().toISOString().split("T")[0]} 12:00`,
    isRead: false
  });

  saveDB(db).catch(console.error);
  auditLog("admin", "ADD", `استبعاد الخادم: ${servant.name} (السبب: ${reason})`);
  res.status(201).json(newExcluded);
});

// -----------------------------------------------------------------
// MESSAGES CENTRE ENDPOINTS
// -----------------------------------------------------------------
app.get("/api/messages", async (req, res) => {
  const db = loadDB();
  res.json(db.messages);
});

app.post("/api/messages", async (req, res) => {
  const { senderId, senderName, title, content, isGeneral, recipientIds } = req.body;
  const db = loadDB();

  if (!title || !content) {
    return res.status(400).json({ error: "العنوان ومضمون الرسالة مطلوبان" });
  }

  const now = new Date();
  const dateStr = `${now.toISOString().split("T")[0]} ${now.toTimeString().substring(0, 5)}`;

  const newMsg: Message = {
    id: `msg-${Date.now()}`,
    senderId: senderId || "u-1",
    senderName: senderName || "الإدارة",
    title,
    content,
    date: dateStr,
    isGeneral: !!isGeneral,
    recipientIds: isGeneral ? undefined : recipientIds || [],
    readBy: []
  };

  db.messages.push(newMsg);

  db.notifications.unshift({
    id: `notif-${Date.now()}-msg`,
    title: "إرسال تعميم/رسالة جديدة",
    content: `تم نشر ${isGeneral ? "تعميم عام لجميع الخدام" : "رسالة خاصة موجهة للخدام"} بعنوان (${title}).`,
    type: "INFO",
    date: `${new Date().toISOString().split("T")[0]} 12:00`,
    isRead: false
  });

  saveDB(db).catch(console.error);
  auditLog("supervisor", "ADD", `إرسال رسالة جديدة بعنوان: ${title}`);
  res.status(201).json(newMsg);
});

app.post("/api/messages/:id/read", async (req, res) => {
  const { id } = req.params;
  const { servantId } = req.body;
  if (!servantId) return res.status(400).json({ error: "معرّف الخادم مطلوب" });

  const db = loadDB();
  const msgIdx = db.messages.findIndex(m => m.id === id);
  if (msgIdx === -1) return res.status(404).json({ error: "الرسالة غير موجودة" });

  if (!db.messages[msgIdx].readBy.includes(servantId)) {
    db.messages[msgIdx].readBy.push(servantId);
    saveDB(db).catch(console.error);
  }

  res.json({ success: true });
});

app.post("/api/messages/delete-all", async (req, res) => {
  const db = loadDB();
  db.messages = [];
  saveDB(db).catch(console.error);
  auditLog("admin", "DELETE", "حذف جميع التعميمات والرسائل الإدارية من النظام");
  res.json({ success: true, message: "تم حذف جميع الرسائل بنجاح" });
});

app.delete("/api/messages/:id", async (req, res) => {
  const { id } = req.params;
  const db = loadDB();
  const idx = db.messages.findIndex(m => m.id === id);
  if (idx === -1) return res.status(404).json({ error: "الرسالة غير موجودة" });

  const msg = db.messages[idx];
  db.messages.splice(idx, 1);
  saveDB(db).catch(console.error);
  auditLog("admin", "DELETE", `حذف الرسالة الإدارية: ${msg.title}`);
  res.json({ success: true, message: "تم حذف الرسالة بنجاح" });
});

// -----------------------------------------------------------------
// NOTIFICATIONS CENTRE ENDPOINTS
// -----------------------------------------------------------------
app.get("/api/notifications", async (req, res) => {
  const db = loadDB();
  res.json(db.notifications);
});

app.post("/api/notifications/mark-read", async (req, res) => {
  const db = loadDB();
  db.notifications.forEach(n => { n.isRead = true; });
  saveDB(db).catch(console.error);
  res.json({ success: true });
});

app.post("/api/notifications/:id/read", async (req, res) => {
  const { id } = req.params;
  const db = loadDB();
  const n = db.notifications.find(item => item.id === id);
  if (n) {
    n.isRead = true;
    saveDB(db).catch(console.error);
  }
  res.json({ success: true });
});

app.post("/api/notifications/dismiss-all", async (req, res) => {
  const db = loadDB();
  db.notifications = [];
  saveDB(db).catch(console.error);
  res.json({ success: true });
});

// -----------------------------------------------------------------
// AUDIT LOGS ENDPOINTS
// -----------------------------------------------------------------
app.get("/api/logs", async (req, res) => {
  const db = loadDB();
  res.json(db.auditLogs);
});

// -----------------------------------------------------------------
// GOOGLE SHEETS SYNC ENDPOINTS
// -----------------------------------------------------------------
app.get("/api/sheets-sync/config", async (req, res) => {
  const db = loadDB();
  res.json({
    config: db.sheetsSyncConfig,
    logs: db.sheetsSyncLogs
  });
});

app.post("/api/sheets-sync/config", async (req, res) => {
  const { spreadsheetId, clientEmail, privateKey, isEnabled } = req.body;
  const db = loadDB();

  const cleanId = (spreadsheetId || "").trim();
  const match = cleanId.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  const finalId = match && match[1] ? match[1] : cleanId;

  db.sheetsSyncConfig = {
    spreadsheetId: finalId,
    clientEmail: (clientEmail || "").trim(),
    privateKey: privateKey || "",
    isEnabled: !!isEnabled,
    lastSyncedAt: db.sheetsSyncConfig?.lastSyncedAt || "",
    lastError: ""
  };

  saveDB(db).catch(console.error);
  auditLog("admin", "SYNC", `تحديث إعدادات ربط Google Sheets (الحالة: ${isEnabled ? "مفعل" : "معطل"})`);
  res.json({ success: true, config: db.sheetsSyncConfig });
});

// RFC 4180 Compliant CSV Parser
function parseCSV(text: string): string[][] {
  const lines: string[][] = [];
  let row: string[] = [];
  let inQuotes = false;
  let currentVal = "";
  
  const normalizedText = text.replace(/\r\n/g, "\n").replace(/\r/g, "");
  
  for (let i = 0; i < normalizedText.length; i++) {
    const char = normalizedText[i];
    const nextChar = normalizedText[i + 1];
    
    if (inQuotes) {
      if (char === '"') {
        if (nextChar === '"') {
          currentVal += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        currentVal += char;
      }
    } else {
      if (char === '"') {
        inQuotes = true;
      } else if (char === ',') {
        row.push(currentVal);
        currentVal = "";
      } else if (char === '\n') {
        row.push(currentVal);
        lines.push(row);
        row = [];
        currentVal = "";
      } else {
        currentVal += char;
      }
    }
  }
  
  if (currentVal || row.length > 0) {
    row.push(currentVal);
    lines.push(row);
  }
  
  return lines;
}
// Function to push full reports and annual statistics to the second tab "التقارير" in Google Sheets
async function pushReportsToGoogleSheet(force: boolean = false, accessToken?: string, customSpreadsheetId?: string) {
  const db = loadDB();
  if (force) {
    try {
      await generateAndPushReportsToGoogleSheet(accessToken);
      console.log("[Sync Reports] Reports pushed to Google Sheets successfully.");
    } catch (e: any) {
      console.error("[Sync Reports Error]", e.message || e);
      throw e;
    }
  } else {
    console.log("[Sync Reports] Google Sheets sync is disabled. Skipping real-time reports push.");
  }
}

// Mappings registry for dynamically linking Google Sheets columns to local settings.
// This allows automatically matching Arabic/English headers and setting properties.
interface SettingField {
  key: string;
  headers: string[];
  getValue: (s: any) => any;
  setValue: (s: any, val: any) => boolean; // returns true if changed
}

const settingsMapping: SettingField[] = [
  {
    key: "headerTitle",
    headers: ["عنوان_المنصة", "عنوان المنصة", "header_title", "headertitle"],
    getValue: (s) => s.headerTitle || "",
    setValue: (s, val) => {
      const v = String(val).trim();
      if (v && s.headerTitle !== v) {
        s.headerTitle = v;
        return true;
      }
      return false;
    }
  },
  {
    key: "headerSub",
    headers: ["عنوان_المنصة_الفرعي", "عنوان المنصة الفرعي", "header_sub", "headersub"],
    getValue: (s) => s.headerSub || "",
    setValue: (s, val) => {
      const v = String(val).trim();
      if (v && s.headerSub !== v) {
        s.headerSub = v;
        return true;
      }
      return false;
    }
  },
  {
    key: "themeColor",
    headers: ["لون_النظام_الأساسي", "لون النظام الأساسي", "theme_color", "themecolor", "اللون_الأساسي", "اللون الأساسي"],
    getValue: (s) => s.theme?.primaryColor || "emerald",
    setValue: (s, val) => {
      const v = String(val).trim();
      if (v) {
        if (!s.theme) s.theme = {};
        if (s.theme.primaryColor !== v) {
          s.theme.primaryColor = v;
          return true;
        }
      }
      return false;
    }
  },
  {
    key: "logoUrl",
    headers: ["رابط_الشعار", "رابط الشعار", "logo_url", "logourl"],
    getValue: (s) => s.logoUrl || "",
    setValue: (s, val) => {
      const v = String(val).trim();
      if (v && !v.includes("[صورة شخصية مرفوعة") && !v.startsWith("[Base64") && s.logoUrl !== v) {
        s.logoUrl = v;
        return true;
      }
      return false;
    }
  },
  {
    key: "cardBackgroundUrl",
    headers: ["رابط_خلفية_البطاقة", "رابط خلفية البطاقة", "card_background_url", "cardbgurl", "خلفية_البطاقة", "خلفية البطاقة"],
    getValue: (s) => s.cardBackgroundUrl || "",
    setValue: (s, val) => {
      const v = String(val).trim();
      if (!v) return false;
      let updated = false;
      if (!v.includes("[صورة شخصية مرفوعة") && s.cardBackgroundUrl !== v) {
        s.cardBackgroundUrl = v;
        updated = true;
      }
      if (!s.card) s.card = {};
      if (s.card.bannerUrl !== v) {
        s.card.bannerUrl = v;
        updated = true;
      }
      return updated;
    }
  },
  {
    key: "cardLogoUrl",
    headers: ["رابط_شعار_خاصة_بالبطاقة", "رابط شعار خاصة بالبطاقة", "card_logo_url", "cardlogourl", "شعار_البطاقة", "شعار البطاقة"],
    getValue: (s) => s.cardLogoUrl || "",
    setValue: (s, val) => {
      const v = String(val).trim();
      if (!v) return false;
      let updated = false;
      if (!v.includes("[صورة شخصية مرفوعة") && s.cardLogoUrl !== v) {
        s.cardLogoUrl = v;
        updated = true;
      }
      if (!s.card) s.card = {};
      if (s.card.logoUrl !== v) {
        s.card.logoUrl = v;
        updated = true;
      }
      return updated;
    }
  },
  {
    key: "logoPositionOnCard",
    headers: ["موضع_الشعار_بالبطاقة", "موضع الشعار بالبطاقة", "موضع_شعار_البطاقة", "موضع شعار البطاقة", "card_logo_position", "cardlogoposition"],
    getValue: (s) => s.card?.logoPosition === "left" ? "يسار" : s.card?.logoPosition === "center" ? "وسط" : "يمين",
    setValue: (s, val) => {
      if (!s.card) s.card = {};
      const v = String(val).trim();
      const pos = (v === "يسار" || v.toLowerCase() === "left") ? "left" : (v === "وسط" || v.toLowerCase() === "center") ? "center" : "right";
      if (s.card.logoPosition !== pos) {
        s.card.logoPosition = pos;
        return true;
      }
      return false;
    }
  },
  {
    key: "logoSizeOnCard",
    headers: ["حجم_الشعار_بالبكسل", "حجم الشعار بالبكسل", "حجم_شعار_البطاقة", "حجم شعار البطاقة", "card_logo_size", "cardlogosize"],
    getValue: (s) => s.card?.logoSize || 45,
    setValue: (s, val) => {
      if (!s.card) s.card = {};
      const num = Number(val);
      if (!isNaN(num) && num > 0 && s.card.logoSize !== num) {
        s.card.logoSize = num;
        return true;
      }
      return false;
    }
  },
  {
    key: "showHeaderBrandingText",
    headers: ["إظهار_نصوص_الترويسة", "اظهار نصوص الترويسة", "إظهار نصوص الترويسة", "show_header_branding", "showheaderbranding"],
    getValue: (s) => s.showHeaderBrandingText !== false ? "نعم" : "لا",
    setValue: (s, val) => {
      const v = String(val).trim();
      const show = (v === "نعم" || v.toLowerCase() === "yes" || v.toLowerCase() === "true" || v === "مفعل" || v === "مفتوح");
      if (s.showHeaderBrandingText !== show) {
        s.showHeaderBrandingText = show;
        return true;
      }
      return false;
    }
  },
  {
    key: "headerTitleSize",
    headers: ["حجم_عنوان_المنصة", "حجم عنوان المنصة", "header_title_size", "headertitlesize"],
    getValue: (s) => s.headerTitleSize || "lg",
    setValue: (s, val) => {
      const v = String(val).trim().toLowerCase();
      if ((v === "sm" || v === "md" || v === "lg" || v === "xl") && s.headerTitleSize !== v) {
        s.headerTitleSize = v as any;
        return true;
      }
      return false;
    }
  },
  {
    key: "headerTitleColor",
    headers: ["لون_عنوان_المنصة", "لون عنوان المنصة", "header_title_color", "headertitlecolor"],
    getValue: (s) => s.headerTitleColor || "#ffffff",
    setValue: (s, val) => {
      const v = String(val).trim();
      if (v && s.headerTitleColor !== v) {
        s.headerTitleColor = v;
        return true;
      }
      return false;
    }
  },
  {
    key: "overlayOpacity",
    headers: ["درجة_تعتيم_الغلاف", "درجة تعتيم الغلاف", "overlay_opacity", "overlayopacity"],
    getValue: (s) => s.overlayOpacity !== undefined ? s.overlayOpacity : 60,
    setValue: (s, val) => {
      const num = Number(val);
      if (!isNaN(num) && num >= 0 && num <= 100 && s.overlayOpacity !== num) {
        s.overlayOpacity = num;
        return true;
      }
      return false;
    }
  },
  {
    key: "headerHeight",
    headers: ["ارتفاع_الترويسة", "ارتفاع الترويسة", "header_height", "headerheight"],
    getValue: (s) => s.headerHeight || 256,
    setValue: (s, val) => {
      const num = Number(val);
      if (!isNaN(num) && num > 0 && s.headerHeight !== num) {
        s.headerHeight = num;
        return true;
      }
      return false;
    }
  },
  {
    key: "coverUrl",
    headers: ["رابط_صورة_الغلاف", "رابط صورة الغلاف", "cover_url", "coverurl", "رابط_الغلاف", "رابط الغلاف", "صورة_الغلاف", "صورة الغلاف"],
    getValue: (s) => s.coverUrl || "",
    setValue: (s, val) => {
      const v = String(val).trim();
      if (s.coverUrl !== v) {
        s.coverUrl = v;
        return true;
      }
      return false;
    }
  },
  {
    key: "tickerIsEnabled",
    headers: ["حالة_شريط_الأخبار", "حالة شريط الأخبار", "ticker_status", "tickerstatus"],
    getValue: (s) => s.ticker?.isEnabled ? "نعم" : "لا",
    setValue: (s, val) => {
      const v = String(val).trim();
      const enabled = (v === "نعم" || v.toLowerCase() === "yes" || v.toLowerCase() === "true" || v === "مفعل" || v === "مفتوح");
      if (!s.ticker) s.ticker = {};
      if (s.ticker.isEnabled !== enabled) {
        s.ticker.isEnabled = enabled;
        return true;
      }
      return false;
    }
  },
  {
    key: "tickerNews",
    headers: ["نص_شريط_الأخبار", "نص شريط الأخبار", "news_ticker", "newsticker"],
    getValue: (s) => Array.isArray(s.ticker?.news) ? s.ticker.news.join(" | ") : (s.ticker?.news || ""),
    setValue: (s, val) => {
      const v = String(val).trim();
      if (!s.ticker) s.ticker = {};
      const currentNews = Array.isArray(s.ticker.news) ? s.ticker.news.join(" | ") : "";
      if (currentNews !== v) {
        s.ticker.news = v ? v.split("|").map((item: string) => item.trim()) : [];
        return true;
      }
      return false;
    }
  }
];

// Function to push settings to Google Sheets "Settings" tab in real-time (instant sync)
async function pushSettingsToGoogleSheet(settings: any, force: boolean = false, accessToken?: string, customSpreadsheetId?: string) {
  if (process.env.ENABLE_GOOGLE_SHEETS_SYNC !== "true") {
    console.log("[Sync Settings] ENABLE_GOOGLE_SHEETS_SYNC is not 'true'. Skipping pushSettingsToGoogleSheet.");
    return;
  }
  
  try {
    const db = loadDB();
    if (!db.sheetsSyncConfig?.isEnabled && !accessToken) {
      console.log("[Sync Settings] Google Sheets sync is disabled. Skipping settings push.");
      return;
    }
    const email = db.sheetsSyncConfig?.clientEmail;
    const pkey = db.sheetsSyncConfig?.privateKey;
    if (!accessToken && (!email || !pkey)) {
       console.log("[Sync Settings] No auth credentials configured. Skipping settings push.");
       return;
    }
    const rawId = resolveSpreadsheetId(customSpreadsheetId);

    if (!rawId) throw new Error("لم يتم إعداد معرف جدول البيانات Google Sheets. يرجى إعداده من الإعدادات قبل المزامنة.");
  const cleanId = rawId.trim();
    const match = cleanId.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    const spreadsheetId = match && match[1] ? match[1] : cleanId;
  console.log("[SYNC_UNIFIED_SPREADSHEET_ID]", spreadsheetId);

    let authClient: any = null;
    if (email && pkey) {
      authClient = new google.auth.JWT({
        email: email,
        key: pkey.replace(/\\n/g, '\n'),
        scopes: ['https://www.googleapis.com/auth/spreadsheets']
      });
      await authClient.authorize();
    } else if (accessToken) {
      const oauth2Client = new google.auth.OAuth2();
      oauth2Client.setCredentials({ access_token: accessToken });
      authClient = oauth2Client;
    }

    if (!authClient) {
      throw new Error("لم يتم العثور على صلاحيات للاتصال بـ Google Sheets. تأكد من إعداد حساب الخدمة أو تسجيل الدخول.");
    }

    const sheets = google.sheets({ version: "v4", auth: authClient });

    // Check if sheet with title "Settings" exists, else create it
    const meta = await withSheetsRetry(() => sheets.spreadsheets.get({ spreadsheetId }));
    const hasSettingsTab = meta.data.sheets?.some(s => s.properties?.title === "Settings");

    if (!hasSettingsTab) {
      console.log("[Sync Settings] 'Settings' tab not found, creating it...");
      await withSheetsRetry(() => sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [
            {
              addSheet: {
                properties: {
                  title: "Settings",
                  gridProperties: { rowCount: 10, columnCount: 30 }
                }
              }
            }
          ]
        }
      }));
    }

    // Try fetching existing headers from Settings sheet to preserve custom column layouts/additions
    let existingHeaders: string[] = [];
    try {
      const getRes = await withSheetsRetry(() => sheets.spreadsheets.values.get({
        spreadsheetId,
        range: "Settings!A1:AZ1"
      }));
      existingHeaders = (getRes.data.values?.[0] || [])
        .map((h: any) => String(h || "").trim())
        .filter(Boolean);
    } catch (e) {
      console.log("[Sync Settings] Could not fetch existing settings headers (Settings tab might be empty).");
    }

    let finalHeaders: string[] = [];
    let finalValues: any[] = [];

    // The primary Arabic headers from our registry in default order
    const standardHeaders = settingsMapping.map(m => m.headers[0]);

    if (existingHeaders.length > 0) {
      // Use existing headers in their exact order to preserve sheet structure
      finalHeaders = [...existingHeaders];
      
      // Populate values for existing headers
      for (const header of finalHeaders) {
        const normHeader = header.toLowerCase().replace(/_/g, " ");
        const field = settingsMapping.find(m => 
          m.headers.some(h => h.toLowerCase().replace(/_/g, " ") === normHeader)
        );
        if (field) {
          const val = field.getValue(settings);
          finalValues.push(val === undefined || val === null ? "" : val);
        } else {
          finalValues.push(""); // Preserve unknown columns
        }
      }

      // Append any standard headers that are currently missing in the sheet so they get added automatically
      for (const stdHeader of standardHeaders) {
        const normStd = stdHeader.toLowerCase().replace(/_/g, " ");
        const exists = finalHeaders.some(h => h.toLowerCase().replace(/_/g, " ") === normStd);
        if (!exists) {
          finalHeaders.push(stdHeader);
          const field = settingsMapping.find(m => m.headers[0] === stdHeader);
          const val = field ? field.getValue(settings) : "";
          finalValues.push(val === undefined || val === null ? "" : val);
        }
      }
    } else {
      // No existing headers - write all standard fields in default order
      for (const field of settingsMapping) {
        finalHeaders.push(field.headers[0]);
        const val = field.getValue(settings);
        finalValues.push(val === undefined || val === null ? "" : val);
      }
    }

    const rows = [finalHeaders, finalValues.map(sanitizeCellForGoogleSheets)];

    // Helper to convert index to A-Z, AA-ZZ column letter
    const getColumnLetter = (colIndex: number): string => {
      let temp = colIndex;
      let letter = "";
      while (temp >= 0) {
        letter = String.fromCharCode((temp % 26) + 65) + letter;
        temp = Math.floor(temp / 26) - 1;
      }
      return letter;
    };

    // Clear previous settings range first to clean up any old values
    await withSheetsRetry(() => sheets.spreadsheets.values.clear({
      spreadsheetId,
      range: "Settings!A1:AZ20"
    }));

    // Write fresh settings
    // Dynamically calculate the range based on headers length
    const lastColIndex = finalHeaders.length - 1;
    const lastColLetter = lastColIndex >= 0 ? getColumnLetter(lastColIndex) : "Z";
    
    await withSheetsRetry(() => sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `Settings!A1:${lastColLetter}2`,
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: rows
      }
    }));

    const nowLog = new Date().toISOString();
    console.log(`[${nowLog}] [WRITE] pushSettingsToGoogleSheet -> Google Sheets`);
  } catch (err: any) {
    if (err.message && err.message.includes("لم يتم العثور على صلاحيات")) {
      // Expected if no Google Sheets credentials configured, suppress error log
    } else {
      console.error("[Sync Settings] Failed to push settings to Google Sheets:", err.message || err);
    }
    throw err;
  }
}

// Function to pull and sync settings from Google Sheets "Settings" tab to the local DB
async function syncSettingsFromGoogleSheet(customAccessToken?: string, force: boolean = false, customSpreadsheetId?: string) {
  if (process.env.ENABLE_GOOGLE_SHEETS_SYNC !== "true") {
    console.log("[Sync Settings] ENABLE_GOOGLE_SHEETS_SYNC is not 'true'. Skipping syncSettingsFromGoogleSheet.");
    return;
  }
  
  const db = loadDB();
  if (!customAccessToken && !db.sheetsSyncConfig?.isEnabled) {
    console.log("[Sync Settings] Google Sheets sync is disabled. Skipping pull settings.");
    return;
  }
  const email = db.sheetsSyncConfig?.clientEmail;
  const pkey = db.sheetsSyncConfig?.privateKey;
  if (!customAccessToken && (!email || !pkey)) {
     console.log("[Sync Settings] No auth credentials configured. Skipping settings sync.");
     return;
  }
  const rawId = resolveSpreadsheetId(customSpreadsheetId);


  if (!rawId) throw new Error("لم يتم إعداد معرف جدول البيانات Google Sheets. يرجى إعداده من الإعدادات قبل المزامنة.");
  const cleanId = rawId.trim();
  const match = cleanId.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  const spreadsheetId = match && match[1] ? match[1] : cleanId;
  console.log("[SYNC_UNIFIED_SPREADSHEET_ID]", spreadsheetId);

  let rows: any[][] | null = null;

  try {
    let authClient: any = null;
    let token = customAccessToken;
    if (token && token !== "null" && token !== "undefined" && token !== "") {
      const oauth2Client = new google.auth.OAuth2();
      oauth2Client.setCredentials({ access_token: token });
      authClient = oauth2Client;
    } else if (email && pkey) {
      authClient = new google.auth.JWT({
        email: email,
        key: pkey.replace(/\\n/g, '\n'),
        scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly']
      });
      await authClient.authorize();
    }

    if (!authClient) {
      console.log("[Sync Settings] No auth token or credentials. Skipping pull settings.");
      return;
    }

    console.log("[TRACE_PULL_STARTED]");
    console.log("[TRACE_SPREADSHEET_ID]", spreadsheetId);
    console.log("[TRACE_ACCESS_TOKEN]", !!customAccessToken);
        console.log("[TRACE_AUTH_CLIENT]", !!authClient);
    if (authClient) console.log("[TRACE_GOOGLE_CONNECTED]");
    const sheets = google.sheets({ version: "v4", auth: authClient });
    // Fetch spreadsheet metadata to check if "Settings" exists
    const meta = await sheets.spreadsheets.get({ spreadsheetId });
    const hasSettingsTab = meta.data.sheets?.some(s => s.properties?.title === "Settings");

    if (!hasSettingsTab) {
      console.log("[Sync Settings] No 'Settings' tab found in Sheets. Pushing current settings as default.");
      // Tab does not exist, let's create and push current settings!
      await pushSettingsToGoogleSheet(db.settings, true, customAccessToken);
      return;
    }

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: "Settings!A1:AZ2", // Read up to 52 columns dynamically to support custom added settings columns!
    });

    rows = response.data.values || [];
  } catch (err: any) {
    console.error("[Sync Settings] Failed to fetch settings from Google Sheets:", err.message || err);
    // Let's also check if public export for CSV works as a fallback!
    try {
      const csvUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=csv&sheet=Settings`;
      const csvRes = await fetch(csvUrl);
      if (csvRes.ok) {
        const text = await csvRes.text();
        if (!text.trim().startsWith("<!")) {
          const parsed = parseCSV(text);
          if (parsed && parsed.length > 0) {
            rows = parsed;
          }
        }
      }
    } catch (csvErr: any) {
      console.warn("[Sync Settings] Public settings CSV export fallback failed too:", csvErr.message || csvErr);
    }
  }

  if (rows && rows.length > 1) {
    const headers = rows[0].map((h: any) => String(h).trim());
    const values = rows[1];

    const currentSettings = db.settings || {
      headerTitle: "هيئة الحجة ابن الحسن",
      headerSub: "(عَجَّلَ اللَّهُ تَعَالَى فَرَجَهُ)",
      cardBackgroundUrl: "",
      showHeaderBrandingText: true,
      theme: { primaryColor: "emerald", mode: "LIGHT" }
    };

    let changed = false;

    // Dynamically match any columns present in row 0 using our mapping registry!
    for (let i = 0; i < headers.length; i++) {
      const header = headers[i];
      if (!header) continue;

      const normHeader = header.toLowerCase().replace(/_/g, " ");
      const field = settingsMapping.find(m => 
        m.headers.some(h => h.toLowerCase().replace(/_/g, " ") === normHeader)
      );

      if (field && values[i] !== undefined) {
        const rawValue = values[i];
        const isUpdated = field.setValue(currentSettings, rawValue);
        if (isUpdated) {
          changed = true;
        }
      }
    }

    if (changed) {
      console.log("[Sync Settings] Settings synchronized and updated from Google Sheets!");
      // db.settings = currentSettings; // Disabled to prevent resetting settings from old Sheets
      saveDB(db).catch(console.error);
    }
  }
}

// Function to generate report table and update "التقارير" tab in Google Sheets
async function generateAndPushReportsToGoogleSheet(accessToken?: string, customSpreadsheetId?: string) {
  return;
}

// Function to push all servant records to Google Sheets "المعلومات" or first tab (manual bulk push)
async function pushAllServantsToGoogleSheet(force: boolean = false, accessToken?: string, customSpreadsheetId?: string) {
  if (process.env.ENABLE_GOOGLE_SHEETS_SYNC !== "true") {
    console.log("[Sync Sheets] ENABLE_GOOGLE_SHEETS_SYNC is not 'true'. Skipping pushAllServantsToGoogleSheet.");
    return;
  }
  
  try {
    const db = loadDB();
    if (!db.sheetsSyncConfig?.isEnabled && !accessToken) {
      console.log("[Sync Servants] Google Sheets sync is disabled. Skipping servants push.");
      return;
    }
    const email = db.sheetsSyncConfig?.clientEmail;
    const pkey = db.sheetsSyncConfig?.privateKey;
    if (!accessToken && (!email || !pkey)) {
       console.log("[Sync Servants] No auth credentials configured. Skipping servants push.");
       return;
    }
    const rawId = resolveSpreadsheetId(customSpreadsheetId);

    if (!rawId) throw new Error("لم يتم إعداد معرف جدول البيانات Google Sheets. يرجى إعداده من الإعدادات قبل المزامنة.");
  const cleanId = rawId.trim();
    const match = cleanId.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    const spreadsheetId = match && match[1] ? match[1] : cleanId;
  console.log("[SYNC_UNIFIED_SPREADSHEET_ID]", spreadsheetId);

    let authClient: any = null;
    if (email && pkey) {
      console.log(`[Trace] Creating Service Account authClient...`);
      authClient = new google.auth.JWT({
        email: email,
        key: pkey.replace(/\\n/g, '\n'),
        scopes: ['https://www.googleapis.com/auth/spreadsheets']
      });
      await authClient.authorize();
      console.log(`[Trace] Service Account authClient created and authorized.`);
    } else if (accessToken) {
      console.log(`[Trace] Received accessToken. Creating OAuth2 authClient...`);
      const oauth2Client = new google.auth.OAuth2();
      oauth2Client.setCredentials({ access_token: accessToken });
      authClient = oauth2Client;
      console.log(`[Trace] OAuth2 authClient created with provided accessToken.`);
    }

    if (!authClient) {
      console.error(`[Trace] No valid accessToken or service account found!`);
      throw new Error("لم يتم العثور على صلاحيات للاتصال بـ Google Sheets. تأكد من إعداد حساب الخدمة أو تسجيل الدخول.");
    }

    console.log(`[Trace] Opening Google Sheets API client...`);
    console.log("[TRACE_PUSH_STARTED]");
    console.log("[TRACE_SPREADSHEET_ID]", spreadsheetId);
    console.log("[TRACE_ACCESS_TOKEN]", !!accessToken);
    console.log("[TRACE_AUTH_CLIENT]", !!authClient);
    if (authClient) console.log("[TRACE_GOOGLE_CONNECTED]");
    const sheets = google.sheets({ version: "v4", auth: authClient });
    // Fetch spreadsheet metadata to get the first sheet's title
    console.log(`[Trace] Fetching spreadsheet metadata for spreadsheetId: ${spreadsheetId}`);
    const meta = await withSheetsRetry(() => sheets.spreadsheets.get({ spreadsheetId }));
    const sheetsList = meta.data.sheets || [];
    let targetSheetTitle = "المعلومات";
    
    // Check if there is an existing sheet named "المعلومات" or "الخدام"
    const foundSheet = sheetsList.find(s => {
      const t = s.properties?.title;
      return t === "المعلومات" || t === "الخدام" || t === "ورقة المعلومات";
    });

    if (foundSheet && foundSheet.properties?.title) {
      targetSheetTitle = foundSheet.properties.title;
    } else if (sheetsList.length > 0 && sheetsList[0].properties?.title) {
      const firstTitle = sheetsList[0].properties.title;
      const trimmedTitle = firstTitle.trim();
      const isDefaultName = [
        "Sheet1", "ورقة1", "الورقة1", "ورقة الورقة1", "ورقة الورقة 1", 
        "ورقة ١", "ورقة 1", "Sheet 1", "الورقة 1", "الورقة ١", 
        "الورقة الأولى", "الورقة الاولى", "الورقة"
      ].includes(trimmedTitle) || sheetsList.length === 1;

      if (isDefaultName) {
        console.log(`[Sync Servants] Renaming default/single sheet '${firstTitle}' to 'المعلومات'...`);
        await withSheetsRetry(() => sheets.spreadsheets.batchUpdate({
          spreadsheetId,
          requestBody: {
            requests: [
              {
                updateSheetProperties: {
                  properties: {
                    sheetId: sheetsList[0].properties?.sheetId,
                    title: "المعلومات"
                  },
                  fields: "title"
                }
              }
            ]
          }
        })).catch(e => console.error("[Sync Servants] Error renaming sheet:", e));
        targetSheetTitle = "المعلومات";
      } else {
        targetSheetTitle = firstTitle;
      }
    }

    const baseHeaders = [
      "الكود", "الاسم", "رقم الهاتف", "تاريخ الانضمام", "منصب المشرف", 
      "حالة إنسانية", "سبب العذر", "تاريخ الميلاد", "العنوان", "ملاحظات", "الحالة",
      "رابط الصورة الشخصية", "موقع ثابت"
    ];

    const customHeaderSet = new Set<string>();
    const servants = db.servants || [];
    for (const s of servants) {
      if (s.customFields) {
        for (const key of Object.keys(s.customFields)) {
          customHeaderSet.add(key);
        }
      }
    }
    const customHeadersList = Array.from(customHeaderSet);
    const headers = [...baseHeaders, ...customHeadersList];

    const rows = [headers];

    for (const s of servants) {
      let arabicStatus = "فعال";
      if (s.status === ServantStatus.ACTIVE) arabicStatus = "فعال";
      else if (s.status === ServantStatus.SUSPENDED_TEMP) arabicStatus = "موقوف مؤقتا";
      else if (s.status === ServantStatus.SUSPENDED) arabicStatus = "موقوف";
      else if (s.status === ServantStatus.EXCLUDED) arabicStatus = "مستبعد";
      else if (s.status === ServantStatus.ON_LEAVE) arabicStatus = "بإجازة";
      else if (s.status === ServantStatus.WARNING) arabicStatus = "تنبيه";

      let pinnedLocationName = "";
      if (s.pinnedLocationId) {
        const loc = db.locations.find((l: any) => l.id === s.pinnedLocationId);
        if (loc) {
          pinnedLocationName = loc.name;
        }
      }

      // Check if avatar is a very long base64 string
      const avatarToPush = s.avatar || "";

      const rowData = [
        s.code || "",
        s.name || "",
        s.phone || "",
        s.joinDate || "",
        s.isSupervisor ? "مشرف" : "خادم",
        s.humanitarian ? "نعم" : "لا",
        s.humanitarianReason || "",
        s.birthDate || "",
        s.address || "",
        s.notes || "",
        arabicStatus,
        avatarToPush,
        pinnedLocationName
      ];

      for (const customHeader of customHeadersList) {
        rowData.push(s.customFields && s.customFields[customHeader] !== undefined ? s.customFields[customHeader] : "");
      }

      rows.push(rowData);
    }

    // Clear target sheet first
    console.log(`[Trace] Preparing to clear data from sheet: '${targetSheetTitle}'`);
    await withSheetsRetry(() => sheets.spreadsheets.values.clear({
      spreadsheetId,
      range: `'${targetSheetTitle}'!A1:Z5000`
    }));

    const getColumnLetter = (colIndex: number): string => {
      let temp = colIndex;
      let letter = "";
      while (temp >= 0) {
        letter = String.fromCharCode((temp % 26) + 65) + letter;
        temp = Math.floor(temp / 26) - 1;
      }
      return letter;
    };

    const lastColLetter = getColumnLetter(headers.length - 1);

    console.log(`[Trace] Writing ${rows.length} rows to sheet: '${targetSheetTitle}'`);
    // Write fresh servant data
    await withSheetsRetry(() => sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `'${targetSheetTitle}'!A1:${lastColLetter}${rows.length}`,
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: rows.map(r => r.map(sanitizeCellForGoogleSheets))
      }
    }));


    console.log(`[Trace] ✅ Write successful.`);
    console.log(`[Sync Servants] Successfully pushed ${rows.length - 1} servants to sheet [${targetSheetTitle}]`);
  } catch (err: any) {
    console.error(`[Trace] ❌ Write failed: ${err.message || err}`);
    console.error("[Sync Servants] Failed to push servants to Google Sheets:", err.message || err);
    throw err;
  }
}

// Function to push a single servant record to Google Sheets in real-time (instant sync)

// Function to delete a servant record from Google Sheets in real-time (instant sync)
async function deleteServantFromGoogleSheet(code: string) {
  // Automatic sync is strictly disabled per user request - only manual sync is allowed
  console.log("[Sync Delete] Real-time automatic servant delete is disabled per user request. Skipping.");
  return;
}

// Unified function to sync google sheets data using either Service Account or User Access Token
async function syncGoogleSheetsData(customAccessToken?: string, username: string = "makzemos@gmail.com", customSpreadsheetId?: string) {
  
  
  console.log("[TRACE_PULL_STARTED]");
  // Will log the variables later when they are defined.

  (global as any).isSyncingFromSheets = true;

  // Sync system settings first (manually triggered, so force=true)
  await syncSettingsFromGoogleSheet(customAccessToken, true, customSpreadsheetId).catch(e => {
    console.error("[Sync Settings Error during run]", e);
  });

  const db = loadDB();
  const rawId = resolveSpreadsheetId(customSpreadsheetId);
  if (!rawId) throw new Error("لم يتم إعداد معرف جدول البيانات Google Sheets. يرجى إعداده من الإعدادات قبل المزامنة.");
  const cleanId = rawId.trim();
  const match = cleanId.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  const spreadsheetId = match && match[1] ? match[1] : cleanId;
  console.log("[SYNC_UNIFIED_SPREADSHEET_ID]", spreadsheetId);
  
  console.log("[REPORT_SPREADSHEET_SOURCE_FETCH]", spreadsheetId);
  
  let rows: any[][] | null = null;
  let firstSheetTitle = "الكشف الرئيسي (CSV)";
  let isPublicSync = false;
  let accessToken = customAccessToken || "";
  let sheetsList: any[] = [];

  const now = new Date();
  const dateStr = now.toISOString().split("T")[0];
  const timeStr = now.toTimeString().split(" ")[0];

  // 1. Primary Attempt: Use OAuth if available, otherwise try CSV
  if (!accessToken && !db.sheetsSyncConfig.clientEmail) {
    // Try public CSV to completely bypass credentials and login prompts if shared as 'Anyone with link'
  try {
    const possibleUrls = [
      `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=csv`,
      `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=csv&gid=0`,
      `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=csv&gid=1867220445`
    ];

    for (const csvUrl of possibleUrls) {
      console.log(`[Sync] Attempting public CSV fetch: ${csvUrl}`);
      const csvRes = await fetch(csvUrl);
      if (csvRes.ok) {
        const text = await csvRes.text();
        const isHtml = text.trim().startsWith("<!") || 
                       text.trim().toLowerCase().startsWith("<html") || 
                       text.includes("ServiceLogin") || 
                       text.includes("google-signin") ||
                       text.includes("<meta") ||
                       text.includes("<script");

        if (isHtml) {
          console.log(`[Sync] CSV export on URL [${csvUrl}] returned HTML (private sheet or incorrect GID).`);
        } else {
          const parsed = parseCSV(text);
          if (parsed && parsed.length > 0 && parsed[0].length > 0) {
            const tempHeaders = parsed[0].map((h: any) => String(h).trim().toLowerCase());
            const hasCode = tempHeaders.some((h: string) => h === "id" || h.includes("كود") || h.includes("code") || h.includes("الرقم") || h.includes("معرف"));
            const hasName = tempHeaders.some((h: string) => h === "الاسم" || h === "الاسم الكامل" || h.includes("اسم") || h.includes("name"));
            
            if (hasCode && hasName) {
              rows = parsed;
              isPublicSync = true;
              console.log(`[Sync] Public CSV download successful via [${csvUrl}]! Parsed ${rows.length} rows.`);
              break; // Found working URL, exit loop
            } else {
              console.log(`[Sync] Public CSV from [${csvUrl}] lacks 'كود' or 'الاسم' headers. Skipping this tab/URL.`);
            }
          }
        }
      } else {
        console.log(`[Sync] CSV export URL [${csvUrl}] returned status ${csvRes.status}`);
      }
    } // end for
  } catch (csvErr: any) {
    console.warn("[Sync] Public CSV fetch failed or was blocked, trying authenticated APIs:", csvErr.message || csvErr);
  }
  

  // 2. Fallback: Authenticated API (requiring OAuth token or Service Account)
  if (!rows) {
    accessToken = customAccessToken || "";
    
    // If no accessToken provided, attempt Service Account authentication
    if (!accessToken) {
      const email = db.sheetsSyncConfig.clientEmail;
      const pkey = db.sheetsSyncConfig.privateKey;
      if (email && pkey) {
        try {
          const jwtClient = new google.auth.JWT({
            email: email,
            key: pkey.replace(/\\n/g, '\n'),
            scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly']
          });
          const credentials = await jwtClient.authorize();
          accessToken = credentials.access_token || undefined;
        } catch (err: any) {
          console.error("Service Account Auth failed:", err);
          throw new Error(`فشل ترخيص حساب الخدمة (Service Account): ${err.message || err}`);
        }
      }
    }
    
    console.log("[TRACE_SPREADSHEET_ID]", spreadsheetId);
    console.log("[TRACE_ACCESS_TOKEN]", !!accessToken);
    console.log("[TRACE_AUTH_CLIENT]", false); // It doesn't use authClient here
    if (accessToken) console.log("[TRACE_GOOGLE_CONNECTED]");
    if (!accessToken) {
      throw new Error("يرجى إما مشاركة جدول البيانات لعامة الناس للقرأة (Anyone with link can view) أو إعداد حساب الخدمة / تسجيل الدخول لتتمكن من المزامنة.");
    }

    // Fetch spreadsheet metadata to get the first sheet's title
    const metaUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}`;
    const metaRes = await fetch(metaUrl, {
      headers: {
        "Authorization": `Bearer ${accessToken}`,
        "Accept": "application/json"
      }
    });

    if (!metaRes.ok) {
      const errorText = await metaRes.text();
      console.error("Google Sheets Metadata Error:", errorText);
      let errorMsg = `فشل جلب بيانات الجدول من Google: ${metaRes.statusText}`;
      if (metaRes.status === 404) {
        errorMsg = "جدول البيانات غير موجود أو ليس لديك صلاحية للوصول إليه (Not Found). يرجى التأكد من صحة معرف الجدول (Spreadsheet ID) ومشاركته مع حساب الخدمة أو تسجيل الدخول بالحساب الصحيح الذي يمتلك صلاحية الوصول.";
      } else if (metaRes.status === 403) {
        errorMsg = "ليس لديك الصلاحية الكافية للوصول لجدول البيانات هذا (Forbidden). يرجى مشاركة الجدول مع البريد الإلكتروني لحساب الخدمة وإعطائه صلاحية 'محرر' (Editor)، أو مشاركة الرابط للجميع كـ 'عارض' (Viewer) للمزامنة اللحظية الفورية بدون أذونات.";
      }
      throw new Error(errorMsg);
    }

    const metaData = await metaRes.json() as any;
    sheetsList = metaData.sheets || [];
    const foundSheet = sheetsList.find((s: any) => {
      const t = s.properties?.title;
      return t === "المعلومات" || t === "الخدام" || t === "ورقة المعلومات";
    });

    if (foundSheet && foundSheet.properties?.title) {
      firstSheetTitle = foundSheet.properties.title;
    } else if (sheetsList.length > 0 && sheetsList[0].properties?.title) {
      firstSheetTitle = sheetsList[0].properties.title;
    } else {
      firstSheetTitle = "Sheet1";
    }

    // Fetch the values of the first sheet
    const valuesUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(firstSheetTitle)}!A1:Z2000`;
    const valuesRes = await fetch(valuesUrl, {
      headers: {
        "Authorization": `Bearer ${accessToken}`,
        "Accept": "application/json"
      }
    });

    if (!valuesRes.ok) {
      let errorMsg = `فشل جلب قيم الخلايا من Google: ${valuesRes.statusText}`;
      if (valuesRes.status === 404) {
        errorMsg = "الصفحة الأولى في جدول البيانات غير موجودة أو معرّف الجدول خاطئ.";
      } else if (valuesRes.status === 403) {
        errorMsg = "صلاحيات غير كافية لجلب خلايا الجدول. تأكد من مشاركة الجدول مع حساب الخدمة أو تسجيل الدخول بحساب مأذون له.";
      }
      throw new Error(errorMsg);
    }

    const valuesData = await valuesRes.json() as any;
    rows = valuesData.values;
  }

  if (!rows || rows.length === 0) {
    throw new Error("جدول البيانات فارغ أو لا يحتوي على صفوف.");
  }

  const headers = rows[0].map((h: any) => String(h).trim().toLowerCase());
  const rawOriginalHeaders = rows[0].map((h: any) => String(h).trim());

  // Header mappings
  const codeIdx = headers.findIndex((h: string) => h === "id" || h.includes("كود") || h.includes("code") || h.includes("الرقم") || h.includes("معرف"));
  const nameIdx = headers.findIndex((h: string) => h === "الاسم" || h === "الاسم الكامل" || h.includes("اسم") || h.includes("name"));
  const phoneIdx = headers.findIndex((h: string) => h.includes("هاتف") || h.includes("جوال") || h.includes("تليفون") || h.includes("phone") || h.includes("mobile"));
  const joinDateIdx = headers.findIndex((h: string) => h.includes("انضمام") || h.includes("تسجيل") || h.includes("join") || (h.includes("تاريخ") && !h.includes("ميلاد")));
  const notesIdx = headers.findIndex((h: string) => h.includes("ملاحظات") || h.includes("ملاحظه") || h.includes("notes") || h.includes("تفاصيل"));
  const supervisorIdx = headers.findIndex((h: string) => h.includes("منصب") || h.includes("المنصب") || h.includes("مشرف") || h.includes("supervisor") || h.includes("إداري"));
  const humanitarianIdx = headers.findIndex((h: string) => h.includes("إنساني") || h.includes("صحي") || h.includes("مرضي") || h.includes("humanitarian"));
  const humanitarianReasonIdx = headers.findIndex((h: string) => h.includes("سبب") || h.includes("السبب") || h.includes("عذر") || h.includes("reason"));
  const birthDateIdx = headers.findIndex((h: string) => h.includes("ميلاد") || h.includes("birth"));
  const addressIdx = headers.findIndex((h: string) => h.includes("سكن") || h.includes("عنوان") || h.includes("address") || h.includes("مقيم"));
  const statusIdx = headers.findIndex((h: string) => h === "الحالة" || h.includes("حالة") || h.includes("status"));
  const pinnedLocationIdx = headers.findIndex((h: string) => h.includes("موقع") || h.includes("مكان") || h.includes("location") || h.includes("ثابت"));
  const accessCodeIdx = headers.findIndex((h: string) => h === "رمز الدخول" || h === "رمز_الدخول" || h.includes("access") || h.includes("رمز_دخول") || h.includes("رمز دخول"));
  const avatarIdx = headers.findIndex((h: string) => h === "رابط الصورة الشخصية" || h === "رابط_الصورة_الشخصية" || h.includes("صورة") || h.includes("صوره") || h.includes("avatar") || h.includes("picture") || h.includes("image"));
  
  if (avatarIdx !== -1) {
    console.log("[PHOTO_COLUMN_FOUND]\nColumn=" + rawOriginalHeaders[avatarIdx]);
  }

  const matchedIndices = new Set([
    codeIdx, nameIdx, phoneIdx, joinDateIdx, notesIdx, supervisorIdx,
    humanitarianIdx, humanitarianReasonIdx, birthDateIdx, addressIdx,
    statusIdx, pinnedLocationIdx, accessCodeIdx, avatarIdx
  ].filter(idx => idx !== -1));

  let addedCount = 0;
  let updatedCount = 0;
  let skippedCount = 0;
  const sheetsCodes = new Set<string>();

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length === 0) continue;

    const rawCodeVal = codeIdx !== -1 ? row[codeIdx] : row[0];
    const rawNameVal = nameIdx !== -1 ? row[nameIdx] : row[2];

    if (!rawCodeVal || !rawNameVal) {
      skippedCount++;
      continue;
    }

    const code = String(rawCodeVal).trim();
    const name = String(rawNameVal).trim();

    if (!code || !name) {
      skippedCount++;
      continue;
    }

    sheetsCodes.add(code);

    const phone = phoneIdx !== -1 && row[phoneIdx] ? String(row[phoneIdx]).trim() : (row[3] ? String(row[3]).trim() : "");
    
    let joinDate = joinDateIdx !== -1 && row[joinDateIdx] ? String(row[joinDateIdx]).trim() : (row[4] ? String(row[4]).trim() : "");
    if (!joinDate || isNaN(Date.parse(joinDate))) {
      joinDate = new Date().toISOString().split("T")[0];
    }

    const positionVal = supervisorIdx !== -1 && row[supervisorIdx] ? String(row[supervisorIdx]).trim().toLowerCase() : (row[5] ? String(row[5]).trim().toLowerCase() : "");
    const isSupervisor = positionVal.includes("مشرف") || positionVal.includes("إداري") || positionVal.includes("نعم") || positionVal === "true" || positionVal === "1";

    const humanitarianVal = humanitarianIdx !== -1 && row[humanitarianIdx] ? String(row[humanitarianIdx]).trim().toLowerCase() : (row[6] ? String(row[6]).trim().toLowerCase() : "");
    const humanitarian = humanitarianVal.includes("نعم") || humanitarianVal === "true" || humanitarianVal === "1" || humanitarianVal.includes("صحي") || humanitarianVal.includes("مرضي");

    const humanitarianReason = humanitarianReasonIdx !== -1 && row[humanitarianReasonIdx] ? String(row[humanitarianReasonIdx]).trim() : (row[7] ? String(row[7]).trim() : "");
    
    const birthDate = birthDateIdx !== -1 && row[birthDateIdx] ? String(row[birthDateIdx]).trim() : (row[8] ? String(row[8]).trim() : "");
    const address = addressIdx !== -1 && row[addressIdx] ? String(row[addressIdx]).trim() : (row[9] ? String(row[9]).trim() : "");
    const notes = notesIdx !== -1 && row[notesIdx] ? String(row[notesIdx]).trim() : (row[10] ? String(row[10]).trim() : "");

    // Parse status from sheet if present
    let status = ServantStatus.ACTIVE;
    if (statusIdx !== -1 && row[statusIdx]) {
      const statusVal = String(row[statusIdx]).trim().toLowerCase();
      if (statusVal === "فعال" || statusVal === "نشط" || statusVal === "active" || statusVal === "نعم") {
        status = ServantStatus.ACTIVE;
      } else if (statusVal === "مجمد" || statusVal === "موقوف مؤقتا" || statusVal === "suspended_temp" || statusVal.includes("مؤقت")) {
        status = ServantStatus.SUSPENDED_TEMP;
      } else if (statusVal === "موقوف" || statusVal === "suspended" || statusVal === "ملغى" || statusVal.includes("موقف")) {
        status = ServantStatus.SUSPENDED;
      } else if (statusVal === "مستبعد" || statusVal === "excluded") {
        status = ServantStatus.EXCLUDED;
      } else if (statusVal === "بإجازة" || statusVal === "اجازة" || statusVal === "on_leave" || statusVal.includes("إجازة")) {
        status = ServantStatus.ON_LEAVE;
      } else if (statusVal === "تنبيه" || statusVal === "warning") {
        status = ServantStatus.WARNING;
      }
    }

    // Parse pinned location from sheet if present
    let pinnedLocationId = undefined;
    if (pinnedLocationIdx !== -1 && row[pinnedLocationIdx]) {
      const locName = String(row[pinnedLocationIdx]).trim();
      const loc = db.locations.find(l => l.name.trim() === locName);
      if (loc) {
        pinnedLocationId = loc.id;
      }
    }

    // Capture all additional unknown columns dynamically!
    const customFields: Record<string, string> = {};
    for (let j = 0; j < row.length; j++) {
      if (!matchedIndices.has(j) && rawOriginalHeaders[j] && row[j] !== undefined) {
        customFields[rawOriginalHeaders[j]] = String(row[j]).trim();
      }
    }

    // Since we no longer use a separate access code, map accessCode directly to the servant's main code
    let accessCode = code;

    const sheetPhotoUrl = avatarIdx !== -1 && row[avatarIdx] ? String(row[avatarIdx]).trim() : "";
    if (sheetPhotoUrl) {
      console.log("[PHOTO_URL_LOADED]\nServant=" + name + "\nURL=" + sheetPhotoUrl);
    }

    const existingIndex = db.servants.findIndex(s => s.code === code);
    if (existingIndex !== -1) {
      // Determine final avatar to save
      let finalPhotoUrl = db.servants[existingIndex].avatar || "";
      if (sheetPhotoUrl && sheetPhotoUrl.startsWith("data:image")) {
         console.log('[PHOTO_LOADED] Google Sheets (loaded full photo from sheet: ' + sheetPhotoUrl.length + ')');
      } else if (sheetPhotoUrl && sheetPhotoUrl.includes("[صورة شخصية مرفوعة")) {
         console.log('[PHOTO_LOADED] Google Sheets (loaded placeholder from sheet: ' + sheetPhotoUrl.length + ')');
      }
      if (finalPhotoUrl && finalPhotoUrl.startsWith("data:image")) {
         console.log('[PHOTO_SAVED] dbInMemory (existing photo found in local DB: ' + finalPhotoUrl.length + ')');
      }

      if (sheetPhotoUrl) {
        const isPlaceholder = sheetPhotoUrl.startsWith("[") && sheetPhotoUrl.endsWith("]");
        const lowerSheetPhotoUrl = sheetPhotoUrl.toLowerCase();
        
        if (lowerSheetPhotoUrl === "حذف" || lowerSheetPhotoUrl === "delete" || lowerSheetPhotoUrl === "remove") {
          console.log('[PHOTO_REMOVED] Source=Google Sheets (Action: DELETE/REMOVE command found)');
          finalPhotoUrl = "";
        } else if (!isPlaceholder) {
          // If it's a real new URL or value, update it
          finalPhotoUrl = sheetPhotoUrl;
        }
      } else {
        // If the sheet cell is empty, we only clear it if the current local avatar is NOT a base64 string
        if (finalPhotoUrl && !finalPhotoUrl.startsWith("data:")) {
          console.log('[PHOTO_REMOVED] Source=Google Sheets (Sheet empty, clearing non-base64 avatar)');
          finalPhotoUrl = "";
        }
      }

      db.servants[existingIndex] = {
        ...db.servants[existingIndex],
        name,
        phone,
        joinDate,
        isSupervisor,
        humanitarian,
        humanitarianReason,
        birthDate,
        address,
        status: (statusIdx !== -1 && row[statusIdx]) ? status : db.servants[existingIndex].status,
        pinnedLocationId: (pinnedLocationIdx !== -1 && row[pinnedLocationIdx]) ? pinnedLocationId : db.servants[existingIndex].pinnedLocationId,
        photoUrl: finalPhotoUrl,
        avatar: finalPhotoUrl,
        notes: notes || db.servants[existingIndex].notes,
        customFields: Object.keys(customFields).length > 0 ? customFields : db.servants[existingIndex].customFields,
        accessCode: accessCode
      };
      updatedCount++;
    } else {
      let finalNewPhotoUrl = "";
      if (sheetPhotoUrl) {
        const isPlaceholder = sheetPhotoUrl.startsWith("[") && sheetPhotoUrl.endsWith("]");
        if (!isPlaceholder && !sheetPhotoUrl.includes("[صورة شخصية مرفوعة")) {
          finalNewPhotoUrl = sheetPhotoUrl;
        }
      }

      db.servants.push({
        id: `s-${code}`,
        code,
        name,
        phone,
        joinDate,
        status,
        isSupervisor,
        humanitarian,
        humanitarianReason,
        birthDate,
        address,
        pinnedLocationId,
        photoUrl: finalNewPhotoUrl || undefined,
        avatar: finalNewPhotoUrl || undefined,
        notes,
        customFields: Object.keys(customFields).length > 0 ? customFields : undefined,
        accessCode: accessCode
      });
      addedCount++;
    }
  }

  if (sheetsCodes.size === 0) {
    throw new Error("لم يتم العثور على أي بيانات خادم صالحة في تبويب المعلومات. تم إلغاء المزامنة لمنع الحذف المفاجئ لقاعدة البيانات المحلية.");
  }

  let deletedCount = 0;
  // Two-way delete: If a servant is in the local database but is missing from Google Sheets, delete them!
  const servantsToDelete = db.servants.filter(s => s.code && !sheetsCodes.has(s.code));
  for (const s of servantsToDelete) {
    const idx = db.servants.findIndex(serv => serv.id === s.id);
    if (idx !== -1) {
      db.servants.splice(idx, 1);
      // Clean related tables
      db.distributions = db.distributions.filter(d => d.servantId !== s.id);
      db.attendance = db.attendance.filter(a => a.servantId !== s.id);
      db.leaves = db.leaves.filter(l => l.servantId !== s.id);
      db.suspensions = db.suspensions.filter(susp => susp.servantId !== s.id);
      db.excluded = db.excluded.filter(ex => ex.servantId !== s.id);
      deletedCount++;
    }
  }

  const syncLogs: string[] = [
    "✓ الاتصال بـ Google API Service ناجح.",
    "✓ تمت المصادقة بنجاح باستخدام رمز المرور المفوض.",
    `✓ الاتصال بجدول البيانات ID: ${spreadsheetId}`,
    `✓ تم استيراد تبويب: "${firstSheetTitle}"`,
    `✓ تم مسح ومعالجة ${rows.length - 1} صفاً من البيانات.`,
    `✓ النتائج: إضافة ${addedCount} خدام جدد، وتحديث ${updatedCount} ملفات حالية، وتخطي ${skippedCount} صفوف غير صالحة، وحذف ${deletedCount} خدام ملغيين من الشيت.`,
    "✓ اكتملت المزامنة ثنائية الاتجاه (Two-Way Sync) بنجاح مع قاعدة البيانات المحلية!"
  ];

  const newLog: SheetsSyncLog = {
    id: `slog-${Date.now()}`,
    date: dateStr,
    time: timeStr,
    username,
    actionType: "مزامنة لحظية ثنائية الاتجاه",
    status: "SUCCESS",
    details: syncLogs.join("\n")
  };

  // --- ALSO SYNC "سجل الاجازات 2026" SHEET IF PRESENT ---
  try {
    let leavesRows: any[][] | null = null;
    const leavesSheetTitle = "سجل الاجازات 2026";
    const hasLeavesTab = sheetsList.some((s: any) => s.properties?.title === leavesSheetTitle);

    if (hasLeavesTab) {
      console.log("[Sync] 'سجل الاجازات 2026' tab found. Fetching leave entries...");
      const lValuesUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(leavesSheetTitle)}!A1:F1000`;
      const lValuesRes = await fetch(lValuesUrl, {
        headers: {
          "Authorization": `Bearer ${accessToken}`,
          "Accept": "application/json"
        }
      });
      if (lValuesRes.ok) {
        const lValuesData = await lValuesRes.json() as any;
        leavesRows = lValuesData.values || [];
      }
    }


    // --- PULL ABSENT, SUSPENDED, EXCLUDED ---
    const absentSheetTitle = "سجل الغيابات";
    const tempSuspSheetTitle = "سجل الموقوفين مؤقتا";
    const suspSheetTitle = "سجل الموقوفين";
    const exclSheetTitle = "سجل المستبعدين";

    const fetchSheetRows = async (title) => {
      const hasTab = sheetsList.some((s) => s.properties?.title === title);
      if (hasTab) {
        console.log(`[Sync] '${title}' tab found. Fetching...`);
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(title)}!A1:G1000`;
        const res = await fetch(url, {
          headers: { "Authorization": `Bearer ${accessToken}`, "Accept": "application/json" }
        });
        if (res.ok) {
          const data = await res.json();
          return data.values || [];
        }
      }
      return null;
    };

    const absentRowsData = await fetchSheetRows(absentSheetTitle);
    const tempSuspRowsData = await fetchSheetRows(tempSuspSheetTitle);
    const suspRowsData = await fetchSheetRows(suspSheetTitle);
    const exclRowsData = await fetchSheetRows(exclSheetTitle);

    if (absentRowsData && absentRowsData.length > 1) {
      db.attendance = db.attendance.filter(a => !(a.id.startsWith("sheet-") && a.isPresent === false));
      for (let i = 1; i < absentRowsData.length; i++) {
        const row = absentRowsData[i];
        if (!row || row.length < 3) continue;
        const code = String(row[0] || "").trim();
        const dateStr = String(row[2] || row[3] || "").trim();
        const notes = String(row[5] || "").trim();
        const supervisor = String(row[4] || "").trim();
        if (!code) continue;
        const servant = db.servants.find(s => s.code === code || s.id === code);
        if (!servant) continue;
        const absD = parseSheetDate(dateStr) || new Date();
        const yyyy = absD.getFullYear();
        const mm = String(absD.getMonth() + 1).padStart(2, '0');
        const dd = String(absD.getDate()).padStart(2, '0');
        db.attendance.push({
           id: `sheet-abs-${Date.now()}-${i}`,
           servantId: servant.id,
           date: `${yyyy}-${mm}-${dd}`,
           isPresent: false,
           isMarked: true,
           supervisorName: supervisor || "من Google Sheets",
           dutyId: "",
           locationId: ""
           // notes: notes (removed)
        });
      }
    }

    db.suspensions = db.suspensions.filter(s => !s.id.startsWith("sheet-"));

    const parseSuspensionRow = (row, i, isTemp) => {
      if (!row || row.length < 3) return;
      const code = String(row[0] || "").trim();
      const startDateStr = String(row[2] || row[3] || "").trim();
      const durationStr = String(row[4] || "").trim();
      const endDateStr = String(row[5] || "").trim();
      const reason = String(row[6] || "").trim();
      if (!code) return;
      const servant = db.servants.find(s => s.code === code || s.id === code);
      if (!servant) return;
      const startD = parseSheetDate(startDateStr) || new Date();
      const durationMonths = parseInt(durationStr) || (isTemp ? 1 : 24);
      
      const yyyy = startD.getFullYear();
      const mm = String(startD.getMonth() + 1).padStart(2, '0');
      const dd = String(startD.getDate()).padStart(2, '0');
      
      let endY = yyyy, endM = parseInt(mm), endD = parseInt(dd);
      if (endDateStr) {
          const endDObj = parseSheetDate(endDateStr);
          if (endDObj) {
             endY = endDObj.getFullYear();
             endM = endDObj.getMonth() + 1;
             endD = endDObj.getDate();
          }
      }
      
      db.suspensions.push({
         id: `sheet-susp-${Date.now()}-${i}`,
         servantId: servant.id,
         servantCode: servant.code,
         servantName: servant.name,
         startDate: `${yyyy}-${mm}-${dd}`,
         endDate: endDateStr ? `${endY}-${String(endM).padStart(2, '0')}-${String(endD).padStart(2, '0')}` : "",
         durationMonths: durationMonths,
         reason: reason || (isTemp ? "تلقائي من Sheet" : "موقوف من Sheet"),
         status: "ACTIVE"
      });
      servant.status = isTemp ? ServantStatus.SUSPENDED_TEMP : ServantStatus.SUSPENDED;
      servant.statusDate = `${yyyy}-${mm}-${dd}`;
    };

    if (tempSuspRowsData && tempSuspRowsData.length > 1) {
      for (let i = 1; i < tempSuspRowsData.length; i++) {
        parseSuspensionRow(tempSuspRowsData[i], i, true);
      }
    }
    if (suspRowsData && suspRowsData.length > 1) {
      for (let i = 1; i < suspRowsData.length; i++) {
        parseSuspensionRow(suspRowsData[i], i + 10000, false);
      }
    }

    if (exclRowsData && exclRowsData.length > 1) {
      db.excluded = db.excluded.filter(e => !e.id.startsWith("sheet-"));
      for (let i = 1; i < exclRowsData.length; i++) {
        const row = exclRowsData[i];
        if (!row || row.length < 3) continue;
        const code = String(row[0] || "").trim();
        const dateStr = String(row[2] || row[3] || "").trim();
        const reason = String(row[4] || "").trim();
        const notes = String(row[5] || "").trim();
        if (!code) continue;
        const servant = db.servants.find(s => s.code === code || s.id === code);
        if (!servant) continue;
        const exD = parseSheetDate(dateStr) || new Date();
        const yyyy = exD.getFullYear();
        const mm = String(exD.getMonth() + 1).padStart(2, '0');
        const dd = String(exD.getDate()).padStart(2, '0');
        db.excluded.push({
           id: `sheet-excl-${Date.now()}-${i}`,
           servantId: servant.id,
           servantCode: servant.code,
           servantName: servant.name,
           date: `${yyyy}-${mm}-${dd}`,
           reason: reason || "استبعاد من Google Sheets",
           notes: notes
        });
        servant.status = ServantStatus.EXCLUDED;
        servant.statusDate = `${yyyy}-${mm}-${dd}`;
      }
    }
    // --- END PULL ---

    if (leavesRows && leavesRows.length > 1) {
      console.log(`[Sync Leaves] Read ${leavesRows.length - 1} leave records.`);
      // Clear existing sheets leaves
      db.leaves = db.leaves.filter(l => !l.id.startsWith("sheet-"));

      const todayBaghdad = getBaghdadDate();

      for (let i = 1; i < leavesRows.length; i++) {
        const row = leavesRows[i];
        if (!row || row.length < 3) continue;
        const code = String(row[0] || "").trim();
        const name = String(row[1] || "").trim();
        const startDateStr = String(row[2] || "").trim();
        const durationStr = String(row[3] || "").trim();
        const endDateStr = String(row[4] || "").trim();
        const details = String(row[5] || "").trim();

        if (!code) continue;

        const servant = db.servants.find(s => s.code === code);
        if (!servant) continue;

        const startD = parseSheetDate(startDateStr);
        let endD = parseSheetDate(endDateStr);
        let durationDays = 1;

        if (startD) {
          if (endD) {
            durationDays = Math.max(1, Math.round((endD.getTime() - startD.getTime()) / (1000 * 60 * 60 * 24)));
          } else {
            const parsedDuration = parseInt(durationStr);
            const days = isNaN(parsedDuration) || parsedDuration <= 0 ? 1 : parsedDuration;
            durationDays = days;
            endD = new Date(startD.getTime() + days * 24 * 60 * 60 * 1000);
          }

          const formattedStart = startD.toISOString().split("T")[0];
          const formattedEnd = endD.toISOString().split("T")[0];

          let status: "ACTIVE" | "COMPLETED" | "PENDING" = "PENDING";
          if (todayBaghdad.getTime() >= startD.getTime() && todayBaghdad.getTime() <= endD.getTime()) {
            status = "ACTIVE";
          } else if (todayBaghdad.getTime() > endD.getTime()) {
            status = "COMPLETED";
          }

          db.leaves.push({
            id: `sheet-${code}-${formattedStart}`,
            servantId: servant.id,
            servantCode: servant.code,
            servantName: servant.name,
            startDate: formattedStart,
            endDate: formattedEnd,
            durationDays: durationDays,
            reason: details || "إجازة رسمية عبر Google Sheets",
            status: status
          });
        }
      }
      syncLogs.push(`✓ تم استيراد تبويب: "${leavesSheetTitle}" وتحديث سجلات الإجازات بنجاح.`);
  }

      // Update our newLog with the new logs
      newLog.details = syncLogs.join("\n");
    } catch (leavesErr: any) {
    console.error("[Sync] Error syncing leaves from Google Sheet:", leavesErr.message || leavesErr);
  }

  db.sheetsSyncLogs.unshift(newLog);
  db.sheetsSyncConfig.lastSyncedAt = `${dateStr} ${timeStr}`;
  db.sheetsSyncConfig.lastError = "";
  db.pendingChanges = {
    addedServants: [],
    updatedServants: [],
    deletedServants: [],
    settingsUpdated: false,
    logoChanged: false,
    coverChanged: false,
    updatedDuties: [],
    recordedAttendance: 0,
    addedMessages: 0,
    leavesUpdated: false,
    firstChangeTime: null,
    lastChangeTime: null
  };

  db.notifications.unshift({
    id: `notif-${Date.now()}-sync`,
    title: "تمت المزامنة بنجاح مع Google Sheets",
    content: `تم مزامنة وتحديث الخدام بنجاح. (مضاف: ${addedCount}، محدث: ${updatedCount}، محذوف: ${deletedCount})`,
    type: "SUCCESS",
    date: `${dateStr} ${timeStr.substring(0, 5)}`,
    isRead: false
  });

  db.servants.sort((a, b) => {
    const codeA = parseInt(a.code) || 0;
    const codeB = parseInt(b.code) || 0;
    return codeA - codeB;
  });
  saveDB(db).catch(console.error);
  auditLog(username, "SYNC", `مزامنة الخدام بنجاح من Google Sheet (مضاف: ${addedCount}، محدث: ${updatedCount})`);

  // Write back servants with their generated/updated accessCode to Google Sheets
  // This satisfies: تنزيل الرموز في الشيت في ورقة المعلومات وفي عامود رمز_الدخول B
  await pushAllServantsToGoogleSheet(true, accessToken, customSpreadsheetId);

  // Automatic reports generation has been separated and disabled during normal sync.
  // It must be triggered manually by the supervisor.
  console.log("[Sync] Normal sync completed. Automatic reports write-back is skipped.");

  return {
    success: true,
    lastSyncedAt: db.sheetsSyncConfig.lastSyncedAt,
    log: newLog,
    addedCount,
    updatedCount,
    skippedCount
  };


}
}
app.post("/api/sheets-sync/run", async (req, res) => {
  
  
  const { username, accessToken, spreadsheetId: customSpreadsheetId } = req.body;
  const db = loadDB();

  const activeUser = username || "makzemos@gmail.com";
  const now = new Date();
  const dateStr = now.toISOString().split("T")[0];
  const timeStr = now.toTimeString().split(" ")[0];

  try {
    const result = await syncGoogleSheetsData(accessToken, activeUser, customSpreadsheetId);
    res.json(result);
  } catch (error: any) {
    console.error("Sheets Sync Server-Side Error:", error);
    const errorLog: SheetsSyncLog = {
      id: `slog-${Date.now()}`,
      date: dateStr,
      time: timeStr,
      username: activeUser,
      actionType: "مزامنة لحظية ثنائية الاتجاه",
      status: "FAILED",
      details: `فشلت المزامنة اللحظية.
السبب: ${error.message || error}`
    };

    db.sheetsSyncLogs.unshift(errorLog);
    db.sheetsSyncConfig.lastError = error.message || String(error);
    saveDB(db).catch(console.error);

    res.status(500).json({
      success: false,
      message: error.message || "حدث خطأ غير متوقع أثناء المزامنة"
    });
  }
});

// Endpoint to fetch pending changes
app.get("/api/sheets-sync/pending-changes", async (req, res) => {
  const db = loadDB();
  const p = db.pendingChanges || {
    addedServants: [],
    updatedServants: [],
    deletedServants: [],
    settingsUpdated: false,
    logoChanged: false,
    coverChanged: false,
    updatedDuties: [],
    recordedAttendance: 0,
    addedMessages: 0,
    leavesUpdated: false,
    firstChangeTime: null,
    lastChangeTime: null
  };

  const hasUnsaved = !!(
    p.addedServants.length > 0 ||
    p.updatedServants.length > 0 ||
    p.deletedServants.length > 0 ||
    p.settingsUpdated ||
    p.logoChanged ||
    p.coverChanged ||
    p.updatedDuties.length > 0 ||
    p.recordedAttendance > 0 ||
    p.addedMessages > 0 ||
    p.leavesUpdated
  );

  res.json({
    success: true,
    pendingChanges: p,
    hasUnsaved
  });
});

// Endpoint to check conflict dynamically
app.get("/api/sheets-sync/check-conflict", async (req, res) => {
  const db = loadDB();
  const email = db.sheetsSyncConfig?.clientEmail;
  const pkey = db.sheetsSyncConfig?.privateKey;
  const rawId = resolveSpreadsheetId();

  if (!rawId) {
    return res.json({ success: true, conflict: false });
  }

  if (!rawId) throw new Error("لم يتم إعداد معرف جدول البيانات Google Sheets. يرجى إعداده من الإعدادات قبل المزامنة.");
  const cleanId = rawId.trim();
  const match = cleanId.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  const spreadsheetId = match && match[1] ? match[1] : cleanId;
  console.log("[SYNC_UNIFIED_SPREADSHEET_ID]", spreadsheetId);

  try {
    let authClient: any = null;
    if (email && pkey) {
      authClient = new google.auth.JWT({
        email: email,
        key: pkey.replace(/\\n/g, '\n'),
        scopes: ['https://www.googleapis.com/auth/drive.readonly']
      });
      await authClient.authorize();
    }

    if (authClient) {
      const sheetModifiedTime = await getSpreadsheetLastModifiedTime(spreadsheetId, authClient);
      if (sheetModifiedTime && db.sheetsSyncConfig.lastSyncedAt) {
        const lastSyncTime = new Date(db.sheetsSyncConfig.lastSyncedAt);
        if (sheetModifiedTime.getTime() > lastSyncTime.getTime() + 5000) {
          return res.json({
            success: true,
            conflict: true,
            sheetModifiedTime: sheetModifiedTime.toISOString(),
            lastSyncedAt: db.sheetsSyncConfig.lastSyncedAt,
            message: "⚠️ تم تعديل جدول البيانات على Google Sheets بواسطة مستخدم آخر منذ آخر عملية جلب قمت بها."
          });
        }
      }
    }
    res.json({ success: true, conflict: false });
  } catch (error: any) {
    res.json({ success: true, conflict: false, error: error.message });
  }
});

// Endpoint to push manual changes to Google Sheets
app.post("/api/sheets-sync/push", async (req, res) => {
  
  
  const { username, accessToken, force } = req.body;
  const db = loadDB();
  const activeUser = username || "admin";
  const now = new Date();
  const dateStr = now.toISOString().split("T")[0];
  const timeStr = now.toTimeString().split(" ")[0];

  const email = db.sheetsSyncConfig?.clientEmail;
  const pkey = db.sheetsSyncConfig?.privateKey;
  const rawId = resolveSpreadsheetId();

  if (!rawId) {
    return res.status(400).json({ success: false, message: "لم يتم تكوين معرّف جدول البيانات في الإعدادات." });
  }

  if (!rawId) throw new Error("لم يتم إعداد معرف جدول البيانات Google Sheets. يرجى إعداده من الإعدادات قبل المزامنة.");
  const cleanId = rawId.trim();
  const match = cleanId.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  const spreadsheetId = match && match[1] ? match[1] : cleanId;
  console.log("[SYNC_UNIFIED_SPREADSHEET_ID]", spreadsheetId);

  try {
    let authClient: any = null;
    if (email && pkey) {
      authClient = new google.auth.JWT({
        email: email,
        key: pkey.replace(/\\n/g, '\n'),
        scopes: ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.readonly']
      });
      await authClient.authorize();
    } else if (accessToken) {
      const oauth2Client = new google.auth.OAuth2();
      oauth2Client.setCredentials({ access_token: accessToken });
      authClient = oauth2Client;
    }

    console.log("[TRACE_PUSH_STARTED]");
    console.log("[TRACE_SPREADSHEET_ID]", spreadsheetId);
    console.log("[TRACE_ACCESS_TOKEN]", !!accessToken);
    console.log("[TRACE_AUTH_CLIENT]", !!authClient);
    if (authClient) console.log("[TRACE_GOOGLE_CONNECTED]");
    if (!authClient) {
      return res.status(401).json({
        success: false,
        message: "لم يتم العثور على صلاحيات للاتصال بـ Google Sheets (رفع البيانات). يرجى تسجيل الدخول بحساب Google أولاً أو تكوين حساب خدمة مفعل."
      });
    }

    // Check Conflict
    if (!force && authClient) {
      const sheetModifiedTime = await getSpreadsheetLastModifiedTime(spreadsheetId, authClient);
      if (sheetModifiedTime && db.sheetsSyncConfig.lastSyncedAt) {
        const lastSyncTime = new Date(db.sheetsSyncConfig.lastSyncedAt);
        if (sheetModifiedTime.getTime() > lastSyncTime.getTime() + 5000) {
          return res.status(409).json({
            success: false,
            conflict: true,
            message: `⚠️ تحذير: تم تعديل جدول البيانات (Google Sheet) بواسطة مستخدم آخر منذ آخر عملية جلب للبيانات قمت بها (آخر جلب كان في ${db.sheetsSyncConfig.lastSyncedAt}). الرفع الآن قد يؤدي إلى حذف تعديلاته أو حدوث تضارب. هل أنت متأكد من رغبتك في الاستمرار والرفع؟`
          });
        }
      }
    }

    console.log("[Push Sync] Proceeding with manual sequential push to Google Sheets...");

    // Run push operations sequentially with a short delay to prevent rate limits
    await pushAllServantsToGoogleSheet(true, accessToken);
    

    if (db.settings) {
      await pushSettingsToGoogleSheet(db.settings, true, accessToken);
      
    }

    await pushReportsToGoogleSheet(true, accessToken);

    // After a successful push, we clear pendingChanges
    db.pendingChanges = {
      addedServants: [],
      updatedServants: [],
      deletedServants: [],
      settingsUpdated: false,
      logoChanged: false,
      coverChanged: false,
      updatedDuties: [],
      recordedAttendance: 0,
      addedMessages: 0,
      leavesUpdated: false,
      firstChangeTime: null,
      lastChangeTime: null
    };

    const newLog: SheetsSyncLog = {
      id: `slog-${Date.now()}`,
      date: dateStr,
      time: timeStr,
      username: activeUser,
      actionType: "رفع البيانات (Push)",
      status: "SUCCESS",
      details: "✓ تم ترحيل كافة البيانات المحلية والخدام والتقارير والإعدادات بنجاح إلى Google Sheets (رفع يدوي كامل).\n✓ تم تصفير قائمة التغييرات المحلية المعلقة."
    };

    db.sheetsSyncLogs.unshift(newLog);
    db.sheetsSyncConfig.lastSyncedAt = `${dateStr} ${timeStr}`;
    db.sheetsSyncConfig.lastError = "";

    db.notifications.unshift({
      id: `notif-${Date.now()}-push`,
      title: "تم رفع البيانات بنجاح إلى Google Sheets",
      content: "تم تحديث أوراق المعلومات والتقارير والإعدادات سحابياً بنجاح.",
      type: "SUCCESS",
      date: `${dateStr} ${timeStr.substring(0, 5)}`,
      isRead: false
    });

    saveDB(db).catch(console.error);
    auditLog(activeUser, "SYNC", "رفع يدوي ناجح لكافة البيانات إلى Google Sheets");

    res.json({
      success: true,
      message: "✓ تم رفع وترحيل كافة البيانات والتعديلات بنجاح إلى Google Sheets وتصفير سجل التغييرات المعلقة!",
      db
    });
  } catch (error: any) {
    console.error("[Push Sync Error]", error);
    const errorLog: SheetsSyncLog = {
      id: `slog-${Date.now()}`,
      date: dateStr,
      time: timeStr,
      username: activeUser,
      actionType: "رفع البيانات (Push)",
      status: "FAILED",
      details: `فشلت عملية رفع البيانات.
السبب: ${error.message || error}`
    };

    db.sheetsSyncLogs.unshift(errorLog);
    db.sheetsSyncConfig.lastError = error.message || String(error);
    saveDB(db).catch(console.error);

    res.status(500).json({
      success: false,
      message: error.message || "فشلت عملية رفع البيانات إلى Google Sheets. يرجى التأكد من الصلاحيات والاتصال."
    });
  }
});

// Dedicated endpoint to generate and push reports manually to "التقارير" sheet
app.post("/api/sheets-sync/generate-reports", async (req, res) => {
  const { accessToken } = req.body;
  const db = loadDB();
  const now = new Date();
  const dateStr = now.toISOString().split("T")[0];
  const timeStr = now.toTimeString().split(" ")[0];

  try {
    await generateAndPushReportsToGoogleSheet(accessToken);

    db.notifications.unshift({
      id: `notif-${Date.now()}-reports-gen`,
      title: "تم تحديث وحساب السجل الدائم يدوياً",
      content: "تمت إعادة حساب وتحديث ورقة 'السجل الدائم' بنجاح في Google Sheets وحذف ورقة التقارير المرفوضة.",
      type: "SUCCESS",
      date: `${dateStr} ${timeStr.substring(0, 5)}`,
      isRead: false
    });
    saveDB(db).catch(console.error);

    auditLog("admin", "SYNC", "تحديث وحساب السجل الدائم يدوياً بنجاح وكتابتها في Google Sheets");
    res.json({ success: true, message: "تم تحديث وحساب السجل الدائم بنجاح وكتابتها في Google Sheets!" });
  } catch (error: any) {
    console.error("Manual Reports Generation Error:", error);
    res.status(500).json({ success: false, error: error.message || "فشل حساب وتحديث التقارير في الشيت" });
  }
});

// Dedicated endpoint to read and fetch data from the "التقارير" tab in Google Sheets
app.get("/api/sheets-sync/fetch-reports", async (req, res) => {
  return res.status(404).json({ success: false, error: "تم إيقاف وحذف مركز التقارير والإحصائيات نهائياً بطلب من الإدارة." });
});

// Dedicated endpoint to read and fetch data from the "السجل الدائم" tab in Google Sheets

async function pushPermanentLogToGoogleSheet(customAccessToken?: string) {
  if (process.env.ENABLE_GOOGLE_SHEETS_SYNC !== "true") {
    console.log("[Sync Sheets] ENABLE_GOOGLE_SHEETS_SYNC is not 'true'. Skipping pushPermanentLogToGoogleSheet.");
    return;
  }
  
  
  const db = loadDB();
  if (!db.sheetsSyncConfig?.isEnabled && !customAccessToken) {
    console.log("[Sync] Google Sheets sync is disabled. Skipping pushPermanentLogToGoogleSheet.");
    return;
  }
  const email = db.sheetsSyncConfig?.clientEmail;
  const pkey = db.sheetsSyncConfig?.privateKey;
  if (!customAccessToken && (!email || !pkey)) {
     console.log("[Sync] No auth credentials configured. Skipping pushPermanentLogToGoogleSheet.");
     return;
  }
  const rawId = resolveSpreadsheetId();
  if (!rawId) throw new Error("لم يتم إعداد معرف جدول البيانات Google Sheets. يرجى إعداده من الإعدادات قبل المزامنة.");
  const cleanId = rawId.trim();
  const match = cleanId.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  const spreadsheetId = match && match[1] ? match[1] : cleanId;
  console.log("[SYNC_UNIFIED_SPREADSHEET_ID]", spreadsheetId);

  let authClient: any = null;
  if (email && pkey) {
    authClient = new google.auth.JWT({
      email: email,
      key: pkey.replace(/\\n/g, '\n'),
      scopes: ['https://www.googleapis.com/auth/spreadsheets']
    });
    await authClient.authorize();
  } else if (customAccessToken) {
    const oauth2Client = new google.auth.OAuth2();
    oauth2Client.setCredentials({ access_token: customAccessToken });
    authClient = oauth2Client;
  }

  if (!authClient) {
    throw new Error("No authentication available (Service Account or Access Token missing)");
  }

  console.log("[TRACE_PUSH_STARTED]");
  console.log("[TRACE_SPREADSHEET_ID]", spreadsheetId);
    console.log("[TRACE_ACCESS_TOKEN]", !!customAccessToken);
  // customAccessToken might be the name here
  console.log("[TRACE_AUTH_CLIENT]", !!authClient);
  if (authClient) console.log("[TRACE_GOOGLE_CONNECTED]");
  const sheets = google.sheets({ version: "v4", auth: authClient });
  const meta = await withSheetsRetry(() => sheets.spreadsheets.get({ spreadsheetId }));
  
  const servants = db.servants || [];
  const attendance = db.attendance || [];
  
  // --- START OF "السجل الدائم" (Permanent Log) Sheet Generation ---
    const permLogTab = meta.data.sheets?.find((s: any) => {
      const t = s.properties?.title;
      return t && t.trim().normalize("NFC") === "السجل الدائم".normalize("NFC");
    });
    const hasPermLogTab = !!permLogTab;
    let actualSheetName = hasPermLogTab ? permLogTab.properties.title : "السجل الدائم";

    let existingSheetData: any[][] = [];

    if (hasPermLogTab) {
      console.log("[ARCHIVE_SHEET_FOUND]\nSheet=السجل الدائم");
      try {
        const response = await withSheetsRetry(() => sheets.spreadsheets.values.get({
          spreadsheetId,
          range: `'${actualSheetName}'!A1:H10000`
        }));
        existingSheetData = response.data.values || [];
      } catch (getErr) {
        console.warn("[Sync Reports] Could not read existing 'السجل الدائم' sheet:", getErr);
      }
    } else {
      console.log("[Sync Reports] 'السجل الدائم' tab not found, creating it...");
      await withSheetsRetry(() => sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [
            {
              addSheet: {
                properties: {
                  title: "السجل الدائم",
                  gridProperties: { rowCount: 1500, columnCount: 10 }
                }
              }
            }
          ]
        }
      }));
      console.log("[ARCHIVE_SHEET_FOUND]\nSheet=السجل الدائم");
    }

    const permHeaders = [
      "كود",
      "الاسم الثلاثي او الرباعي",
      "حضور",
      "مجاز",
      "غياب",
      "موقوف مؤقت",
      "موقوف",
      "مستبعد"
    ];

    const existingMap = new Map<string, any[]>();
    const existingByNameMap = new Map<string, any[]>();
    for (let i = 1; i < existingSheetData.length; i++) {
      const row = existingSheetData[i];
      if (row) {
        const code = String(row[0] || "").trim();
        const name = String(row[1] || "").trim();
        if (code) {
          existingMap.set(code, row);
        }
        if (name) {
          existingByNameMap.set(name, row);
        }
      }
    }

    const permRows: any[][] = [permHeaders];
    const leaves = db.leaves || [];
    const suspensions = db.suspensions || [];
    const excluded = db.excluded || [];

    const processedCodes = new Set<string>();
    const processedNames = new Set<string>();

    console.log("[ARCHIVE_PUSH_STARTED]");

    for (const servant of servants) {
      const sCode = String(servant.code || "").trim();
      const sName = String(servant.name || "").trim();

      // 1. حضور
      const presentDatesList = attendance
        .filter(a => 
          String(a.servantId) === String(servant.id) && 
          a.isPresent === true && 
          a.isMarked === true
        )
        .map(a => a.date || "")
        .filter(Boolean);
      const uniquePresentDates = Array.from(new Set(presentDatesList));
      const presentValue = uniquePresentDates.join(", ") || "";

      // 2. مجاز
      const servantLeaves = leaves.filter(l => 
        String(l.servantId) === String(servant.id) && 
        l.status !== "REJECTED"
      );
      const leaveRanges = servantLeaves.map(l => {
        if (l.startDate && l.endDate && l.startDate !== l.endDate) {
          return `${l.startDate} إلى ${l.endDate}`;
        }
        return l.startDate || l.endDate || "مجاز";
      });
      if (leaveRanges.length === 0 && servant.status === ServantStatus.ON_LEAVE) {
        leaveRanges.push("مجاز");
      }
      const leavesValue = leaveRanges.join(", ");

      // 3. غياب (مع فحص الإجازات الذكي)
      const absentDatesList = attendance
        .filter(a => String(a.servantId) === String(servant.id) && a.isPresent === false && a.isMarked === true)
        .map(a => a.date || "")
        .filter(Boolean);
      
      const filteredAbsentDates = absentDatesList.filter(dStr => {
        const absD = parseSheetDate(dStr);
        if (!absD) return true;
        // Check if falls within any leave range
        const isOnLeave = servantLeaves.some(l => {
          const startD = parseSheetDate(l.startDate);
          const endD = parseSheetDate(l.endDate);
          if (startD && endD) {
            return absD.getTime() >= startD.getTime() && absD.getTime() <= endD.getTime();
          }
          return false;
        });
        return !isOnLeave;
      });
      const uniqueAbsentDates = Array.from(new Set(filteredAbsentDates));
      const absentValue = uniqueAbsentDates.join(", ") || "";

      // 4. موقوف مؤقت
      const tempSuspEntries = suspensions.filter(s => 
        String(s.servantId) === String(servant.id) && 
        (s.durationMonths <= 1 || s.reason?.includes("تلقائي") || s.reason?.includes("مؤقت"))
      );
      const tempSuspRanges = tempSuspEntries.map(s => {
        if (s.startDate && s.endDate && s.startDate !== s.endDate) {
          return `${s.startDate} إلى ${s.endDate}`;
        }
        return s.startDate || "موقوف مؤقت";
      });
      if (tempSuspRanges.length === 0 && servant.status === ServantStatus.SUSPENDED_TEMP) {
        tempSuspRanges.push(servant.autoSuspensionStartDate || "موقوف مؤقت");
      }
      const tempSuspValue = tempSuspRanges.join(", ");

      // 5. موقوف
      const regSuspEntries = suspensions.filter(s => 
        String(s.servantId) === String(servant.id) && 
        !(s.durationMonths <= 1 || s.reason?.includes("تلقائي") || s.reason?.includes("مؤقت"))
      );
      const regSuspRanges = regSuspEntries.map(s => {
        if (s.startDate && s.endDate && s.startDate !== s.endDate) {
          return `${s.startDate} إلى ${s.endDate}`;
        }
        return s.startDate || "موقوف";
      });
      if (regSuspRanges.length === 0 && servant.status === ServantStatus.SUSPENDED) {
        regSuspRanges.push("موقوف");
      }
      const regSuspValue = regSuspRanges.join(", ");

      // 6. مستبعد
      const exclEntries = excluded.filter(ex => 
        String(ex.servantId) === String(servant.id)
      );
      const exclRanges = exclEntries.map(ex => ex.date || "مستبعد");
      if (exclRanges.length === 0 && servant.status === ServantStatus.EXCLUDED) {
        exclRanges.push("مستبعد");
      }
      const exclValue = exclRanges.join(", ");

      const statusSummary = `حضور: ${presentValue || "بلا"}، مجاز: ${leavesValue || "بلا"}، غياب: ${absentValue || "بلا"}، موقوف مؤقت: ${tempSuspValue || "بلا"}، موقوف: ${regSuspValue || "بلا"}، مستبعد: ${exclValue || "بلا"}`;
      console.log(`[ARCHIVE_RECORD]\nServant=${servant.name}\nStatus=${statusSummary}`);

      // Look up if they have an existing row in Google Sheets
      const existingRow = (sCode ? existingMap.get(sCode) : null) || (sName ? existingByNameMap.get(sName) : null);
      
      let finalPresent = presentValue;
      let finalLeaves = leavesValue;
      let finalAbsent = absentValue;
      let finalTempSusp = tempSuspValue;
      let finalRegSusp = regSuspValue;
      let finalExcl = exclValue;
      
      if (existingRow) {
        finalPresent = mergeCommaValues(existingRow[2], presentValue);
        finalLeaves = mergeCommaValues(existingRow[3], leavesValue);
        finalAbsent = mergeCommaValues(existingRow[4], absentValue);
        finalTempSusp = mergeCommaValues(existingRow[5], tempSuspValue);
        finalRegSusp = mergeCommaValues(existingRow[6], regSuspValue);
        finalExcl = mergeCommaValues(existingRow[7], exclValue);
      }

      const hasAnyStatus = !!(finalPresent || finalLeaves || finalAbsent || finalTempSusp || finalRegSusp || finalExcl);
      if (hasAnyStatus) {
        permRows.push([
          sCode || (existingRow ? existingRow[0] : ""),
          sName || (existingRow ? existingRow[1] : ""),
          finalPresent,
          finalLeaves,
          finalAbsent,
          finalTempSusp,
          finalRegSusp,
          finalExcl
        ]);

        if (sCode) processedCodes.add(sCode);
        if (sName) processedNames.add(sName);
      }
    }

    // Add any remaining historical rows for servants who were deleted from the system entirely
    for (let i = 1; i < existingSheetData.length; i++) {
      const row = existingSheetData[i];
      if (!row || row.length === 0) continue;
      const code = String(row[0] || "").trim();
      const name = String(row[1] || "").trim();
      
      const alreadyProcessed = (code && processedCodes.has(code)) || (name && processedNames.has(name));
      if (!alreadyProcessed) {
        // Pads the row to make sure it has exactly 8 columns (to avoid errors)
        const paddedRow = [...row];
        while (paddedRow.length < 8) paddedRow.push("");
        permRows.push(paddedRow);
      }
    }

    console.log(`[ARCHIVE_RECORD_COUNT]\nCount=${permRows.length - 1} (including preserved deleted/historical records)`);

    // Clear previous "السجل الدائم" so no stale data remains, then write merged data
    await withSheetsRetry(() => sheets.spreadsheets.values.clear({
      spreadsheetId,
      range: `'${actualSheetName}'!A1:H10000`
    }));

    const updateResponse = await withSheetsRetry(() => sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `'${actualSheetName}'!A1:H10000`,
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: permRows.map(r => r.map(sanitizeCellForGoogleSheets))
      }
    }));

    console.log(`[REPORT_SPREADSHEET_ID]\n${spreadsheetId}`);
    console.log(`[PERM_LOG_WRITE_SUCCESS rows=${permRows.length}]`);
    console.log(`[REPORT_TARGET_SHEET]\nالسجل الدائم`);
    console.log(`[REPORT_FIRST_5_ROWS]\n${JSON.stringify(permRows.slice(0, 5), null, 2)}`);
    console.log(`[REPORT_GOOGLE_API_RESPONSE]\n${JSON.stringify(updateResponse.data || updateResponse, null, 2)}`);

    const updatedRows = (updateResponse.data && updateResponse.data.updatedRows) ? updateResponse.data.updatedRows : 0;
    const updatedCells = (updateResponse.data && updateResponse.data.updatedCells) ? updateResponse.data.updatedCells : 0;

    if (updatedRows === 0 || updatedCells === 0) {
      console.log("[GOOGLE_DID_NOT_WRITE]");
    } else {
      console.log("[REPORT_EXPORT_SUCCESS]");
    }

    console.log("[ARCHIVE_WRITE_SUCCESS]");
    console.log(`[Sync Reports] Successfully exported/updated ${permRows.length - 1} permanent log rows in 'السجل الدائم' sheet.`);
    console.log("[ARCHIVE_WRITE]");
    console.log("Sheet=السجل الدائم");
    console.log(`RowsWritten=${permRows.length}`);

    // --- START OF NEW STATUS LOG SHEETS GENERATION ---
    const existingSheetTitles = (meta.data.sheets || []).map((s: any) => s.properties?.title?.trim().normalize("NFC") || "");

    const leaveHeaders = ["كود الخادم", "اسم الخادم", "تاريخ الحالة", "تاريخ البدء", "تاريخ الانتهاء", "المدة بالأيام", "حالة الإجازة", "السبب / الملاحظات"];
    const leaveRows: any[][] = [leaveHeaders];
    const servantsOnLeaveSet = new Set<string>();

    for (const l of leaves) {
      if (l.status === "REJECTED") continue;
      const sCode = String(l.servantCode || "").trim();
      const sName = String(l.servantName || "").trim();
      let statusAr = "قيد الانتظار";
      if (l.status === "ACTIVE") statusAr = "سارية";
      else if (l.status === "COMPLETED") statusAr = "منتهية";

      leaveRows.push([
        sCode,
        sName,
        l.startDate || "",
        l.startDate || "",
        l.endDate || "",
        l.durationDays || "",
        statusAr,
        l.reason || ""
      ]);
      if (l.servantId) servantsOnLeaveSet.add(String(l.servantId));
    }

    for (const s of servants) {
      if (s.status === ServantStatus.ON_LEAVE && !servantsOnLeaveSet.has(String(s.id))) {
        leaveRows.push([
          s.code || "",
          s.name || "",
          s.statusDate || s.joinDate || "",
          "",
          "",
          "",
          "بإجازة (حالة الخادم)",
          ""
        ]);
      }
    }

    const absentHeaders = ["كود الخادم", "اسم الخادم", "تاريخ الحالة", "تاريخ الغياب", "اسم المشرف المسجل", "ملاحظات"];
    const absentRows: any[][] = [absentHeaders];

    const absentAttendance = attendance
      .filter(a => a.isPresent === false && a.isMarked === true)
      .sort((a, b) => new Date(b.date || "").getTime() - new Date(a.date || "").getTime());

    for (const a of absentAttendance) {
      const servant = servants.find(s => String(s.id) === String(a.servantId));
      if (!servant) continue;

      const absD = parseSheetDate(a.date || "");
      let isOnLeave = false;
      if (absD) {
        const servantLeaves = leaves.filter(l => String(l.servantId) === String(servant.id) && l.status !== "REJECTED");
        isOnLeave = servantLeaves.some(l => {
          const startD = parseSheetDate(l.startDate);
          const endD = parseSheetDate(l.endDate);
          if (!startD || !endD) return false;
          return absD.getTime() >= startD.getTime() && absD.getTime() <= endD.getTime();
        });
      }

      if (isOnLeave) continue;

      absentRows.push([
        servant.code || "",
        servant.name || "",
        a.date || "",
        a.date || "",
        a.supervisorName || "النظام",
        a.time ? `سجل غياب في الساعة ${a.time}` : ""
      ]);
    }

    const tempSuspHeaders = ["كود الخادم", "اسم الخادم", "تاريخ الحالة", "تاريخ البدء", "المدة بالأشهر", "تاريخ الانتهاء", "السبب / الملاحظات"];
    const tempSuspRows: any[][] = [tempSuspHeaders];
    const tempSuspServantIds = new Set<string>();

    for (const s of suspensions) {
      const isTemp = (s.durationMonths <= 1 || s.reason?.includes("تلقائي") || s.reason?.includes("مؤقت"));
      if (!isTemp) continue;

      tempSuspRows.push([
        s.servantCode || "",
        s.servantName || "",
        s.startDate || "",
        s.startDate || "",
        s.durationMonths || "",
        s.endDate || "",
        s.reason || ""
      ]);
      if (s.servantId) tempSuspServantIds.add(String(s.servantId));
    }

    for (const s of servants) {
      if (s.status === ServantStatus.SUSPENDED_TEMP && !tempSuspServantIds.has(String(s.id))) {
        tempSuspRows.push([
          s.code || "",
          s.name || "",
          s.autoSuspensionStartDate || s.statusDate || s.joinDate || "",
          s.autoSuspensionStartDate || "",
          "",
          "",
          "موقوف مؤقت (حالة الخادم)"
        ]);
      }
    }

    const regSuspHeaders = ["كود الخادم", "اسم الخادم", "تاريخ الحالة", "تاريخ البدء", "المدة بالأشهر", "تاريخ الانتهاء", "السبب / الملاحظات"];
    const regSuspRows: any[][] = [regSuspHeaders];
    const regSuspServantIds = new Set<string>();

    for (const s of suspensions) {
      const isTemp = (s.durationMonths <= 1 || s.reason?.includes("تلقائي") || s.reason?.includes("مؤقت"));
      if (isTemp) continue;

      regSuspRows.push([
        s.servantCode || "",
        s.servantName || "",
        s.startDate || "",
        s.startDate || "",
        s.durationMonths || "",
        s.endDate || "",
        s.reason || ""
      ]);
      if (s.servantId) regSuspServantIds.add(String(s.servantId));
    }

    for (const s of servants) {
      if (s.status === ServantStatus.SUSPENDED && !regSuspServantIds.has(String(s.id))) {
        regSuspRows.push([
          s.code || "",
          s.name || "",
          s.statusDate || s.joinDate || "",
          "",
          "",
          "",
          "موقوف (حالة الخادم)"
        ]);
      }
    }

    const exclHeaders = ["كود الخادم", "اسم الخادم", "تاريخ الحالة", "التاريخ", "السبب", "ملاحظات إضافية"];
    const exclRows: any[][] = [exclHeaders];
    const exclServantIds = new Set<string>();

    for (const ex of excluded) {
      exclRows.push([
        ex.servantCode || "",
        ex.servantName || "",
        ex.date || "",
        ex.date || "",
        ex.reason || "",
        ex.notes || ""
      ]);
      if (ex.servantId) exclServantIds.add(String(ex.servantId));
    }

    for (const s of servants) {
      if (s.status === ServantStatus.EXCLUDED && !exclServantIds.has(String(s.id))) {
        exclRows.push([
          s.code || "",
          s.name || "",
          s.statusDate || s.joinDate || "",
          "",
          "مستبعد (حالة الخادم)",
          ""
        ]);
      }
    }

    const targetSheets = [
      { title: "سجل الاجازات", rows: leaveRows, keyIndex: [0, 2] },
      { title: "سجل الغيابات", rows: absentRows, keyIndex: [0, 2] },
      { title: "سجل الموقوفين مؤقتا", rows: tempSuspRows, keyIndex: [0, 2] },
      { title: "سجل الموقوفين", rows: regSuspRows, keyIndex: [0, 2] },
      { title: "سجل المستبعدين", rows: exclRows, keyIndex: [0, 2] }
    ];

    const addSheetRequests: any[] = [];
    for (const tSheet of targetSheets) {
      const normTitle = tSheet.title.normalize("NFC");
      if (!existingSheetTitles.includes(normTitle)) {
        addSheetRequests.push({
          addSheet: {
            properties: {
              title: tSheet.title,
              gridProperties: { rowCount: 1500, columnCount: 10 }
            }
          }
        });
      }
    }

    if (addSheetRequests.length > 0) {
      console.log(`[Sync Reports] Creating ${addSheetRequests.length} missing status log sheets...`);
      await withSheetsRetry(() => sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: addSheetRequests
        }
      }));
    }

    for (const tSheet of targetSheets) {
      const actualSheetName = tSheet.title;
      const dataRows = tSheet.rows;
      const keyIndex = tSheet.keyIndex;

      let existingRows: any[][] = [];
      const normTitle = actualSheetName.normalize("NFC");
      
      if (existingSheetTitles.includes(normTitle)) {
        try {
          const lResponse = await withSheetsRetry(() => sheets.spreadsheets.values.get({
            spreadsheetId,
            range: `'${actualSheetName}'!A1:J10000`
          }));
          existingRows = lResponse.data.values || [];
        } catch (readErr) {
          console.warn(`[Sync Reports] Could not read existing '${actualSheetName}' sheet:`, readErr);
        }
      }

      // Merge existing history with newly calculated rows
      const mergedRows = mergeSheetRows(existingRows, dataRows, keyIndex);

      await withSheetsRetry(() => sheets.spreadsheets.values.clear({
        spreadsheetId,
        range: `'${actualSheetName}'!A1:J10000`
      }));

      await withSheetsRetry(() => sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `'${actualSheetName}'!A1:J10000`,
        valueInputOption: "USER_ENTERED",
        requestBody: {
          values: mergedRows.map(r => r.map(sanitizeCellForGoogleSheets))
        }
      }));
      console.log(`[Sync Reports] Successfully written ${mergedRows.length - 1} rows to '${actualSheetName}' (merged with existing history).`);
    }
    // --- END OF NEW STATUS LOG SHEETS GENERATION ---
}

// Function to push the Financial Ledger to Google Sheets
async function pushFinancialLedgerToGoogleSheet(customAccessToken?: string) {
  if (process.env.ENABLE_GOOGLE_SHEETS_SYNC !== "true") {
    console.log("[Sync Sheets] ENABLE_GOOGLE_SHEETS_SYNC is not 'true'. Skipping pushFinancialLedgerToGoogleSheet.");
    return;
  }
  const db = loadDB();
  if (!db.sheetsSyncConfig?.isEnabled && !customAccessToken) {
    console.log("[Sync] Google Sheets sync is disabled. Skipping pushFinancialLedgerToGoogleSheet.");
    return;
  }
  const email = db.sheetsSyncConfig?.clientEmail;
  const pkey = db.sheetsSyncConfig?.privateKey;
  if (!customAccessToken && (!email || !pkey)) {
     console.log("[Sync] No auth credentials configured. Skipping pushFinancialLedgerToGoogleSheet.");
     return;
  }
  const rawId = resolveSpreadsheetId();
  if (!rawId) throw new Error("لم يتم إعداد معرف جدول البيانات Google Sheets. يرجى إعداده من الإعدادات قبل المزامنة.");
  const cleanId = rawId.trim();
  const match = cleanId.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  const spreadsheetId = match && match[1] ? match[1] : cleanId;

  let authClient: any = null;
  if (email && pkey) {
    authClient = new google.auth.JWT({
      email: email,
      key: pkey.replace(/\\n/g, '\n'),
      scopes: ['https://www.googleapis.com/auth/spreadsheets']
    });
    await authClient.authorize();
  } else if (customAccessToken) {
    const oauth2Client = new google.auth.OAuth2();
    oauth2Client.setCredentials({ access_token: customAccessToken });
    authClient = oauth2Client;
  }

  if (!authClient) {
    throw new Error("No authentication available (Service Account or Access Token missing)");
  }

  const sheets = google.sheets({ version: "v4", auth: authClient });
  const meta = await withSheetsRetry(() => sheets.spreadsheets.get({ spreadsheetId }));
  
  const financeRecords = db.financeRecords || [];

  const financeTab = meta.data.sheets?.find((s: any) => {
    const t = s.properties?.title;
    return t && t.trim().normalize("NFC") === "السجل المالي".normalize("NFC");
  });
  const hasFinanceTab = !!financeTab;
  let actualSheetName = hasFinanceTab ? financeTab.properties.title : "السجل المالي";

  if (!hasFinanceTab) {
    console.log("[Sync Finance] 'السجل المالي' tab not found, creating it...");
    await withSheetsRetry(() => sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            addSheet: {
              properties: {
                title: "السجل المالي",
                gridProperties: { rowCount: 5000, columnCount: 12 }
              }
            }
          }
        ]
      }
    }));
  }

  // Clear existing content to avoid leftovers
  await withSheetsRetry(() => sheets.spreadsheets.values.clear({
    spreadsheetId,
    range: `'${actualSheetName}'!A1:L10000`
  }));

  // Build rows
  const headers = ["الكود", "اسم الخادم", "رقم الهاتف", "نوع العملية المالية", "المبلغ", "حالة الدفع", "الملاحظات", "تاريخ العملية", "المشرف المسؤول"];
  const rows = [headers];

  const translateType = (type?: string): string => {
    switch (type) {
      case "SUBSCRIPTION": return "اشتراك شهري";
      case "DONATION": return "تبرع";
      case "SPONSORSHIP": return "كفالة إنسانية";
      case "SPECIAL_FEES": return "رسوم خاصة";
      case "EXPENSE": return "مصروف";
      case "OTHER": return "أخرى";
      default: return type || "";
    }
  };

  const translateStatus = (status?: string): string => {
    switch (status) {
      case "PAID": return "مسدد";
      case "UNPAID": return "غير مسدد";
      case "DEFERRED": return "مؤجل";
      case "EXEMPT": return "معفى";
      case "REMARKS": return "عليه ملاحظات";
      case "NOT_PAID": return "غير واصل";
      default: return status || "";
    }
  };

  financeRecords.forEach((rec) => {
    rows.push([
      rec.servantCode || "",
      rec.servantName || "",
      rec.phone || "",
      translateType(rec.type),
      String(rec.amount || 0),
      translateStatus(rec.status),
      rec.notes || "",
      rec.date || "",
      rec.supervisorName || ""
    ]);
  });

  await withSheetsRetry(() => sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `'${actualSheetName}'!A1:I10000`,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: rows.map(r => r.map(sanitizeCellForGoogleSheets))
    }
  }));

  console.log(`[Sync Finance] Successfully written ${rows.length - 1} rows to '${actualSheetName}'.`);
}

// Function to push the Honoring Records to Google Sheets
async function pushHonoringRecordsToGoogleSheet(customAccessToken?: string) {
  if (process.env.ENABLE_GOOGLE_SHEETS_SYNC !== "true") {
    console.log("[Sync Sheets] ENABLE_GOOGLE_SHEETS_SYNC is not 'true'. Skipping pushHonoringRecordsToGoogleSheet.");
    return;
  }
  const db = loadDB();
  if (!db.sheetsSyncConfig?.isEnabled && !customAccessToken) {
    console.log("[Sync] Google Sheets sync is disabled. Skipping pushHonoringRecordsToGoogleSheet.");
    return;
  }
  const email = db.sheetsSyncConfig?.clientEmail;
  const pkey = db.sheetsSyncConfig?.privateKey;
  if (!customAccessToken && (!email || !pkey)) {
     console.log("[Sync] No auth credentials configured. Skipping pushHonoringRecordsToGoogleSheet.");
     return;
  }
  const rawId = resolveSpreadsheetId();
  if (!rawId) throw new Error("لم يتم إعداد معرف جدول البيانات Google Sheets. يرجى إعداده من الإعدادات قبل المزامنة.");
  const cleanId = rawId.trim();
  const match = cleanId.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  const spreadsheetId = match && match[1] ? match[1] : cleanId;

  let authClient: any = null;
  if (email && pkey) {
    authClient = new google.auth.JWT({
      email: email,
      key: pkey.replace(/\\n/g, '\n'),
      scopes: ['https://www.googleapis.com/auth/spreadsheets']
    });
    await authClient.authorize();
  } else if (customAccessToken) {
    const oauth2Client = new google.auth.OAuth2();
    oauth2Client.setCredentials({ access_token: customAccessToken });
    authClient = oauth2Client;
  }

  if (!authClient) {
    throw new Error("No authentication available (Service Account or Access Token missing)");
  }

  const sheets = google.sheets({ version: "v4", auth: authClient });
  const meta = await withSheetsRetry(() => sheets.spreadsheets.get({ spreadsheetId }));
  
  const honoringRecords = db.honoringRecords || [];

  const honoringTab = meta.data.sheets?.find((s: any) => {
    const t = s.properties?.title;
    return t && t.trim().normalize("NFC") === "المحصلون والتكريم".normalize("NFC");
  });
  const hasHonoringTab = !!honoringTab;
  let actualSheetName = hasHonoringTab ? honoringTab.properties.title : "المحصلون والتكريم";

  if (!hasHonoringTab) {
    console.log("[Sync Honoring] 'المحصلون والتكريم' tab not found, creating it...");
    await withSheetsRetry(() => sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            addSheet: {
              properties: {
                title: "المحصلون والتكريم",
                gridProperties: { rowCount: 5000, columnCount: 12 }
              }
            }
          }
        ]
      }
    }));
  }

  // Clear existing content to avoid leftovers
  await withSheetsRetry(() => sheets.spreadsheets.values.clear({
    spreadsheetId,
    range: `'${actualSheetName}'!A1:L10000`
  }));

  // Build rows
  const headers = ["اسم المحصل", "نوع التكريم", "الجهة المانحة", "وصف التكريم", "التاريخ", "الملاحظات", "إضافة المشرف", "تاريخ الإضافة بالنظام"];
  const rows = [headers];

  const translateType = (type?: string): string => {
    if (type === "THANKS_LETTER") return "كتاب شكر";
    if (type === "GIFT") return "هدية";
    if (type === "OTHER") return "تكريم آخر";
    return type || "";
  };

  for (const record of honoringRecords) {
    const r = [
      record.collectorName || "",
      translateType(record.honoringType),
      record.donor || "",
      record.description || "",
      record.date || "",
      record.notes || "",
      record.supervisorName || "",
      new Date(record.createdAt || Date.now()).toLocaleString("ar-IQ")
    ];
    rows.push(r);
  }

  // Write new content
  await withSheetsRetry(() => sheets.spreadsheets.values.update({
    spreadsheetId,
    range: `'${actualSheetName}'!A1:H${rows.length}`,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: rows.map(r => r.map(sanitizeCellForGoogleSheets))
    }
  }));

  console.log(`[Sync Honoring] Successfully written ${rows.length - 1} rows to '${actualSheetName}'.`);
}
// Endpoint to push manual changes to Google Sheets
app.post("/api/sheets-sync/push-all", async (req, res) => {
  const { username, accessToken } = req.body;
  const activeUser = username || "admin";
  const db = loadDB();

  try {
    const email = db.sheetsSyncConfig?.clientEmail;
    const pkey = db.sheetsSyncConfig?.privateKey;

    if (!email && !pkey && !accessToken) {
      return res.status(401).json({ success: false, message: "صلاحيات غير كافية لرفع البيانات. يرجى تسجيل الدخول أو توفير حساب خدمة." });
    }
    
    if (!db.sheetsSyncConfig?.spreadsheetId) {
      return res.status(400).json({ success: false, message: "لم يتم إعداد معرف جدول البيانات" });
    }

    const now = new Date();
    const dateStr = now.toISOString().split("T")[0];
    const timeStr = now.toTimeString().split(" ")[0];

    let syncLogs = [
      "✓ بدأ رفع البيانات إلى الأرشيف (Google Sheets).",
      `✓ الاتصال بجدول البيانات ID: ${db.sheetsSyncConfig.spreadsheetId}`
    ];

    // Push all servants
    await pushAllServantsToGoogleSheet(true, accessToken);
    syncLogs.push("✓ تم رفع وتحديث ورقة المعلومات بنجاح.");

    // Push permanent log
    await pushPermanentLogToGoogleSheet(accessToken);
    syncLogs.push("✓ تم رفع وتحديث ورقة السجل الدائم بنجاح.");

    // Push financial ledger
    await pushFinancialLedgerToGoogleSheet(accessToken);
    await pushHonoringRecordsToGoogleSheet(accessToken);
    syncLogs.push("✓ تم رفع وتحديث ورقة المحصلون والتكريم بنجاح.");
    syncLogs.push("✓ تم رفع وتحديث ورقة السجل المالي بنجاح.");

    const newLog = {
      id: `slog-${Date.now()}`,
      date: dateStr,
      time: timeStr,
      username: activeUser,
      actionType: "رفع بيانات الأرشيف (Push)",
      status: "SUCCESS" as const,
      details: syncLogs.join("\n")
    };

    db.sheetsSyncLogs = db.sheetsSyncLogs || [];
    db.sheetsSyncLogs.unshift(newLog);
    if (db.sheetsSyncLogs.length > 50) db.sheetsSyncLogs.length = 50;

    db.sheetsSyncConfig.lastSyncedAt = `${dateStr} ${timeStr}`;
    db.sheetsSyncConfig.lastError = "";

    saveDB(db).catch(console.error);

    res.json({ success: true, message: "تم رفع البيانات إلى الأرشيف بنجاح", logs: db.sheetsSyncLogs, config: db.sheetsSyncConfig });
  } catch (err: any) {
    console.error("[Sync Push All] Error:", err);
    
    const now = new Date();
    const dateStr = now.toISOString().split("T")[0];
    const timeStr = now.toTimeString().split(" ")[0];
    const newLog = {
      id: `slog-${Date.now()}`,
      date: dateStr,
      time: timeStr,
      username: activeUser,
      actionType: "رفع بيانات الأرشيف (Push)",
      status: "FAILED" as const,
      details: "فشل تصدير الأرشيف: " + (err.message || err)
    };
    db.sheetsSyncLogs = db.sheetsSyncLogs || [];
    db.sheetsSyncLogs.unshift(newLog);
    if (db.sheetsSyncLogs.length > 50) db.sheetsSyncLogs.length = 50;
    saveDB(db).catch(console.error);

    res.status(500).json({ success: false, message: "فشل ترحيل وتصدير البيانات: " + (err.message || err) });
  }
});

// Dedicated endpoint to push the permanent log independently
app.post("/api/sheets-sync/push-permanent-log", async (req, res) => {
  const { accessToken } = req.body;
  try {
    await pushPermanentLogToGoogleSheet(accessToken);
    return res.json({ success: true, message: "تم تحديث السجل الدائم بنجاح" });
  } catch (err: any) {
    console.error("[Push Permanent Log Error]", err.message || err);
    return res.status(500).json({ success: false, error: err.message || String(err) });
  }
});

app.get("/api/sheets-sync/fetch-permanent-log", async (req, res) => {
  const db = loadDB();
  const rawId = resolveSpreadsheetId();
  const email = db.sheetsSyncConfig?.clientEmail;
  const pkey = db.sheetsSyncConfig?.privateKey;
  const accessToken = (req.query.accessToken as string) || (req.headers.authorization?.startsWith("Bearer ") ? req.headers.authorization.substring(7) : undefined);

  if (!rawId) {
    return res.status(400).json({ success: false, error: "لم يتم تكوين معرّف جدول البيانات في الإعدادات." });
  }

  if (!rawId) throw new Error("لم يتم إعداد معرف جدول البيانات Google Sheets. يرجى إعداده من الإعدادات قبل المزامنة.");
  const cleanId = rawId.trim();
  const match = cleanId.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  const spreadsheetId = match && match[1] ? match[1] : cleanId;
  console.log("[SYNC_UNIFIED_SPREADSHEET_ID]", spreadsheetId);

  console.log("[PERM_LOG_FETCH_START]");

  let rows: any[][] | null = null;
  let errorReasons: string[] = [];
  
  if (!email || !pkey) {
    if (accessToken) {
      try {
        const oauth2Client = new google.auth.OAuth2();
        oauth2Client.setCredentials({ access_token: accessToken });
        const sheets = google.sheets({ version: "v4", auth: oauth2Client });
        
        const meta = await withSheetsRetry(() => sheets.spreadsheets.get({ spreadsheetId }));
        const permLogSheet = meta.data.sheets?.find((s: any) => {
          const t = s.properties?.title;
          return t && t.trim().normalize("NFC") === "السجل الدائم".normalize("NFC");
        });

        let actualSheetName = "السجل الدائم";
        if (!permLogSheet) {
          errorReasons.push("'السجل الدائم' tab not found in metadata. Creating it...");
          await withSheetsRetry(() => sheets.spreadsheets.batchUpdate({
            spreadsheetId,
            requestBody: {
              requests: [
                {
                  addSheet: {
                    properties: {
                      title: "السجل الدائم",
                      gridProperties: { rowCount: 1500, columnCount: 10 }
                    }
                  }
                }
              ]
            }
          }));
        } else {
          actualSheetName = permLogSheet.properties.title;
          console.log(`[PERM_LOG_SHEET_FOUND gid=${permLogSheet.properties.sheetId}]`);
        }

        const valuesRes = await withSheetsRetry(() => sheets.spreadsheets.values.get({
          spreadsheetId,
          range: `'${actualSheetName}'!A1:Z10000`
        }));
        
        rows = valuesRes.data.values || null;
        console.log(`[PERM_LOG_ROWS_READ count=${rows ? rows.length : 0}]`);
      } catch (err: any) {
        console.error("[Sync Fetch Permanent Log] OAuth API fetch failed:", err);
        errorReasons.push(err.message || String(err));
      }
    } else {
      errorReasons.push("clientEmail or privateKey is not configured, and no accessToken provided.");
    }
  } else {
    try {
      const authClient = new google.auth.JWT({
        email: email,
        key: pkey.replace(/\\n/g, '\n'),
        scopes: ['https://www.googleapis.com/auth/spreadsheets']
      });
      await authClient.authorize();
      
      console.log("[TRACE_PUSH_STARTED]");
  console.log("[TRACE_SPREADSHEET_ID]", spreadsheetId);
    console.log("[TRACE_ACCESS_TOKEN]", !!accessToken);
  // customAccessToken might be the name here
  console.log("[TRACE_AUTH_CLIENT]", !!authClient);
  if (authClient) console.log("[TRACE_GOOGLE_CONNECTED]");
  const sheets = google.sheets({ version: "v4", auth: authClient });
  const meta = await withSheetsRetry(() => sheets.spreadsheets.get({ spreadsheetId }));
      const permLogSheet = meta.data.sheets?.find((s: any) => {
        const t = s.properties?.title;
        return t && t.trim().normalize("NFC") === "السجل الدائم".normalize("NFC");
      });

      let actualSheetName = "السجل الدائم";
      if (!permLogSheet) {
        errorReasons.push("'السجل الدائم' tab not found in metadata. Creating it...");
        await withSheetsRetry(() => sheets.spreadsheets.batchUpdate({
          spreadsheetId,
          requestBody: {
            requests: [
              {
                addSheet: {
                  properties: {
                    title: "السجل الدائم",
                    gridProperties: { rowCount: 1500, columnCount: 10 }
                  }
                }
              }
            ]
          }
        }));
      } else {
        actualSheetName = permLogSheet.properties.title;
        console.log(`[PERM_LOG_SHEET_FOUND gid=${permLogSheet.properties.sheetId}]`);
      }

      const valuesRes = await withSheetsRetry(() => sheets.spreadsheets.values.get({
        spreadsheetId,
        range: `'${actualSheetName}'!A1:Z10000`
      }));
      
      rows = valuesRes.data.values || null;
      console.log(`[PERM_LOG_ROWS_READ count=${rows ? rows.length : 0}]`);
      
    } catch (err: any) {
      console.error("[Sync Fetch Permanent Log] Authenticated API fetch failed:", err);
      errorReasons.push(err.message || String(err));
    }
  }

  if (!rows || rows.length === 0) {
    if (db.permanentLog && db.permanentLog.length > 0) {
      console.log(`[Sync Fetch Permanent Log] Fallback to local permanent log database cache. Row count: ${db.permanentLog.length}`);
      const headers = ["كود", "الاسم الثلاثي او الرباعي", "حضور", "مجاز", "غياب", "موقوف مؤقت", "موقوف", "مستبعد"];
      return res.json({ success: true, headers, data: db.permanentLog, reasons: errorReasons });
    }
    return res.status(404).json({ 
      success: false, 
      error: "لا توجد بيانات في ورقة 'السجل الدائم' أو يتعذر الوصول إلى الجدول حالياً. يرجى التأكد من كتابة السجلات أولاً أو تفعيل صلاحيات حساب الخدمة.",
      reasons: errorReasons
    });
  }

  // Parse headers & data dynamically
  const headers = rows[0].map((h: any) => String(h).trim());
  const logData = rows.slice(1).map((row: any[]) => {
    return {
      code: row[0] ? String(row[0]).trim() : "",
      name: row[1] ? String(row[1]).trim() : "",
      present: row[2] ? String(row[2]).trim() : "",
      leaves: row[3] ? String(row[3]).trim() : "",
      absent: row[4] ? String(row[4]).trim() : "",
      tempSuspended: row[5] ? String(row[5]).trim() : "",
      suspended: row[6] ? String(row[6]).trim() : "",
      excluded: row[7] ? String(row[7]).trim() : ""
    };
  }).filter((item: any) => item.code || item.name);

  // Cache in local DB
  db.permanentLog = logData;
  saveDB(db).catch(console.error);

  auditLog("system", "SYNC", `جلب بيانات ورقة السجل الدائم بنجاح من Google Sheets عدد السطور: ${logData.length}`);
  return res.json({ success: true, headers, data: logData, reasons: errorReasons });
});

// Dedicated endpoint to clear the permanent log
app.post("/api/sheets-sync/clear-permanent-log", async (req, res) => {
  const { code, supervisorName } = req.body;
  const db = loadDB();

  if (!code) {
    return res.status(400).json({ success: false, error: "رمز تأكيد المشرف مطلوب لإتمام العملية" });
  }

  const cleanInputCode = code.trim();

  // Validate Code: Must be Admin code ("1972" or "197200") or belong to a supervisor (isSupervisor = true)
  const isAdmin = cleanInputCode === "1972" || cleanInputCode === "197200";
  const supervisor = db.servants.find(s => 
    s.isSupervisor && s.code && String(s.code).trim() === cleanInputCode
  );

  if (!isAdmin && !supervisor) {
    return res.status(403).json({ success: false, error: "كود الدخول المدخل ليس لمشرف أو مدير عام معتمد! لا يمكن تصفير السجل." });
  }

  const actorName = isAdmin ? "محمد زيدان عباس" : (supervisor ? supervisor.name : supervisorName || "مشرف معتمد");

  const rawId = resolveSpreadsheetId();
  const email = db.sheetsSyncConfig?.clientEmail;
  const pkey = db.sheetsSyncConfig?.privateKey;

  if (!rawId) {
    return res.status(400).json({ success: false, error: "لم يتم تكوين معرّف جدول البيانات في الإعدادات." });
  }

  if (!rawId) throw new Error("لم يتم إعداد معرف جدول البيانات Google Sheets. يرجى إعداده من الإعدادات قبل المزامنة.");
  const cleanId = rawId.trim();
  const match = cleanId.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  const spreadsheetId = match && match[1] ? match[1] : cleanId;
  console.log("[SYNC_UNIFIED_SPREADSHEET_ID]", spreadsheetId);

  try {
    let authClient: any = null;
    if (email && pkey) {
      authClient = new google.auth.JWT({
        email: email,
        key: pkey.replace(/\\n/g, '\n'),
        scopes: ['https://www.googleapis.com/auth/spreadsheets']
      });
      await authClient.authorize();
    }

    if (!authClient) {
      return res.status(400).json({ success: false, error: "يتطلب تصفير السجل في شيت جوجل تكوين حساب خدمة (Service Account) فعال في الإعدادات." });
    }

    const sheets = google.sheets({ version: "v4", auth: authClient });
    
    // Clear the permanent log rows except headers (A2 to J10000)
    await sheets.spreadsheets.values.clear({
      spreadsheetId,
      range: "'السجل الدائم'!A2:J10000"
    });

    auditLog(actorName, "DELETE", `تم تصفير وحذف جميع سطور السجل الدائم يدوياً من الشيت بواسطة ${actorName}`);
    
    return res.json({ success: true, message: `تم تصفير السجل الدائم بنجاح بواسطة المشرف: ${actorName}` });
  } catch (error: any) {
    console.error("Clear Permanent Log Error:", error);
    return res.status(500).json({ success: false, error: error.message || "فشل تصفير السجل الدائم في الشيت" });
  }
});

// Dedicated endpoint to import and merge permanent log from Excel file
app.post("/api/sheets-sync/import-permanent-log", async (req, res) => {
  const { importedRows, username } = req.body;
  if (!importedRows || !Array.isArray(importedRows)) {
    return res.status(400).json({ success: false, error: "بيانات الاستيراد غير صالحة أو مفقودة." });
  }

  const db = loadDB();
  const actorName = username || "مشرف معتمد";

  // Helper to merge archive values
  const mergeArchiveValuesLocal = (existingStr: any, newStr: any): string => {
    const eStr = String(existingStr || "").trim();
    const nStr = String(newStr || "").trim();
    if (!eStr) return nStr;
    if (!nStr) return eStr;

    const eItems = eStr.split(/[،,]\s*/).map(x => x.trim()).filter(Boolean);
    const nItems = nStr.split(/[،,]\s*/).map(x => x.trim()).filter(Boolean);

    const unique = Array.from(new Set([...eItems, ...nItems]));
    unique.sort((a, b) => a.localeCompare(b, "ar"));

    return unique.join("، ");
  };

  // 1. Get existing permanent log items from local database
  const currentLogs = db.permanentLog || [];
  const mergedMap = new Map<string, any>();
  const mergedByNameMap = new Map<string, any>();

  // Load existing into maps
  currentLogs.forEach(item => {
    const codeKey = String(item.code || "").trim();
    const nameKey = String(item.name || "").trim();
    if (codeKey) mergedMap.set(codeKey, item);
    if (nameKey) mergedByNameMap.set(nameKey, item);
  });

  let updatedCount = 0;
  let addedCount = 0;

  // 2. Iterate and merge incoming imported rows
  importedRows.forEach(row => {
    const incomingCode = String(row.code || "").trim();
    const incomingName = String(row.name || "").trim();
    if (!incomingCode && !incomingName) return;

    // Find if existing exists
    let existing = (incomingCode && mergedMap.get(incomingCode)) || (incomingName && mergedByNameMap.get(incomingName));

    if (existing) {
      // Merge values
      existing.present = mergeArchiveValuesLocal(existing.present, row.present);
      existing.leaves = mergeArchiveValuesLocal(existing.leaves, row.leaves);
      existing.absent = mergeArchiveValuesLocal(existing.absent, row.absent);
      existing.tempSuspended = mergeArchiveValuesLocal(existing.tempSuspended, row.tempSuspended);
      existing.suspended = mergeArchiveValuesLocal(existing.suspended, row.suspended);
      existing.excluded = mergeArchiveValuesLocal(existing.excluded, row.excluded);
      updatedCount++;
    } else {
      // Add new
      const newItem = {
        code: incomingCode,
        name: incomingName,
        present: String(row.present || "").trim(),
        leaves: String(row.leaves || "").trim(),
        absent: String(row.absent || "").trim(),
        tempSuspended: String(row.tempSuspended || "").trim(),
        suspended: String(row.suspended || "").trim(),
        excluded: String(row.excluded || "").trim(),
      };
      if (incomingCode) mergedMap.set(incomingCode, newItem);
      if (incomingName) mergedByNameMap.set(incomingName, newItem);
      addedCount++;
    }
  });

  // Get full combined array
  const finalLogData: any[] = [];
  const processedCodes = new Set<string>();
  const processedNames = new Set<string>();

  // Collect unique items
  mergedMap.forEach((val, key) => {
    finalLogData.push(val);
    processedCodes.add(key);
    if (val.name) processedNames.add(String(val.name).trim());
  });

  mergedByNameMap.forEach((val, key) => {
    const code = String(val.code || "").trim();
    const alreadyProcessed = (code && processedCodes.has(code)) || processedNames.has(key);
    if (!alreadyProcessed) {
      finalLogData.push(val);
    }
  });

  // Save to db.json / local state
  db.permanentLog = finalLogData;
  saveDB(db).catch(console.error);

  // 3. Write back to Google Sheets if configured
  const rawId = resolveSpreadsheetId();
  const email = db.sheetsSyncConfig?.clientEmail;
  const pkey = db.sheetsSyncConfig?.privateKey;

  let sheetSyncSuccess = false;
  let sheetSyncError = "";

  if (rawId && email && pkey) {
    try {
      if (!rawId) throw new Error("لم يتم إعداد معرف جدول البيانات Google Sheets. يرجى إعداده من الإعدادات قبل المزامنة.");
  const cleanId = rawId.trim();
      const match = cleanId.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
      const spreadsheetId = match && match[1] ? match[1] : cleanId;
  console.log("[SYNC_UNIFIED_SPREADSHEET_ID]", spreadsheetId);

      const authClient = new google.auth.JWT({
        email: email,
        key: pkey.replace(/\\n/g, '\n'),
        scopes: ['https://www.googleapis.com/auth/spreadsheets']
      });
      await authClient.authorize();

      const sheets = google.sheets({ version: "v4", auth: authClient });

      // Build rows with headers
      const permHeaders = [
        "كود",
        "الاسم الثلاثي او الرباعي",
        "حضور",
        "مجاز",
        "غياب",
        "موقوف مؤقت",
        "موقوف",
        "مستبعد"
      ];

      const sheetRows: any[][] = [permHeaders];
      finalLogData.forEach(item => {
        sheetRows.push([
          item.code || "",
          item.name || "",
          item.present || "",
          item.leaves || "",
          item.absent || "",
          item.tempSuspended || "",
          item.suspended || "",
          item.excluded || ""
        ]);
      });

      // Clear existing and write back
      await sheets.spreadsheets.values.clear({
        spreadsheetId,
        range: "'السجل الدائم'!A1:J10000"
      });

      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `'السجل الدائم'!A1:H${sheetRows.length}`,
        valueInputOption: "USER_ENTERED",
        requestBody: {
          values: sheetRows.map(r => r.map(sanitizeCellForGoogleSheets))
        }
      });

      sheetSyncSuccess = true;
    } catch (err: any) {
      console.error("[Import Permanent Log] Sheets sync error:", err);
      sheetSyncError = err.message || String(err);
    }
  } else {
    sheetSyncError = "جدول بيانات Google Sheets غير مكون بالكامل (بريد الخدمة أو المفتاح الخاص مفقود). تم حفظ السجلات محلياً في التطبيق فقط.";
  }

  // Audit log
  auditLog(actorName, "SYNC", `استيراد ودمج السجل الدائم من ملف إكسل (تم دمج وتحديث: ${updatedCount}، تم إضافة: ${addedCount})`);

  // System notification
  db.notifications.unshift({
    id: `notif-${Date.now()}-import-perm`,
    title: "تم استيراد السجل الدائم بنجاح",
    content: `تم معالجة ودمج ملف إكسل بنجاح. مضاف جديد: ${addedCount}، مدمج ومحدث: ${updatedCount}. ${sheetSyncSuccess ? "تمت المزامنة بنجاح مع شيت جوجل." : "تنبيه: " + sheetSyncError}`,
    type: sheetSyncSuccess ? "SUCCESS" : "WARNING",
    date: new Date().toISOString().replace("T", " ").substring(0, 16),
    isRead: false
  });
  saveDB(db).catch(console.error);

  return res.json({
    success: true,
    message: `تمت عملية الاستيراد والدمج بنجاح! تم إضافة ${addedCount} سجل جديد وتحديث ${updatedCount} سجل موجود مسبقاً.`,
    addedCount,
    updatedCount,
    sheetSyncSuccess,
    sheetSyncError
  });
});

// Incoming webhook for real-time sheets updates/deletions (Sheets -> System)
app.post("/api/sheets-sync/webhook", async (req, res) => {
  const { action, code, name, phone, joinDate, notes, status, humanitarian, humanitarianReason, isSupervisor, birthDate, address } = req.body;

  if (!code) {
    return res.status(400).json({ error: "كود الخادم مطلوب لتحديد السجل (Servant code is required)" });
  }

  const db = loadDB();
  const idx = db.servants.findIndex(s => String(s.code).trim() === String(code).trim());

  if (action === "delete") {
    if (idx === -1) {
      return res.status(404).json({ error: "الخادم المطلوب حذفه غير موجود حالياً بالنظام" });
    }
    const servant = db.servants[idx];
    const id = servant.id;
    db.servants.splice(idx, 1);

    // Clean related records in other tables
    db.distributions = db.distributions.filter(d => d.servantId !== id);
    db.attendance = db.attendance.filter(a => a.servantId !== id);
    db.leaves = db.leaves.filter(l => l.servantId !== id);
    db.suspensions = db.suspensions.filter(s => s.servantId !== id);
    db.excluded = db.excluded.filter(e => e.servantId !== id);

    db.notifications.unshift({
      id: `notif-${Date.now()}`,
      title: "حذف خادم عبر مزامنة Google Sheets",
      content: `تم إزالة ملف الخادم ${servant.name} (كود: ${servant.code}) بناءً على حذفه من جدول البيانات.`,
      type: "WARNING",
      date: `${new Date().toISOString().split("T")[0]} 12:00`,
      isRead: false
    });

    saveDB(db).catch(console.error);
    auditLog("system", "DELETE", `حذف خادم تلقائياً عبر مزامنة الشيت: ${servant.name} (كود: ${servant.code})`);
    return res.json({ success: true, message: "تم حذف الخادم وتحديث النظام بنجاح" });
  }

  // Handle update/insert (upsert)
  const isNew = idx === -1;
  const servantStatus = (status === "بإجازة" || status === "ON_LEAVE") ? ServantStatus.ON_LEAVE 
                     : (status === "موقوف" || status === "SUSPENDED") ? ServantStatus.SUSPENDED
                     : (status === "مستبعد" || status === "EXCLUDED") ? ServantStatus.EXCLUDED
                     : ServantStatus.ACTIVE;

  const isHumanitarian = humanitarian === "نعم" || humanitarian === true || String(humanitarian).toLowerCase() === "yes";
  const isSuper = isSupervisor === "مشرف" || isSupervisor === true || String(isSupervisor).toLowerCase() === "supervisor" || String(isSupervisor).includes("مشرف");

  if (isNew) {
    // Insert new servant
    const newServant: Servant = {
      id: `servant-${Date.now()}`,
      code: String(code).trim(),
      name: (name || `خادم جديد ${code}`).trim(),
      phone: phone || "",
      joinDate: joinDate || new Date().toISOString().split("T")[0],
      notes: notes || "",
      status: servantStatus,
      humanitarian: isHumanitarian,
      humanitarianReason: humanitarianReason || "",
      isSupervisor: isSuper,
      birthDate: birthDate || "",
      address: address || ""
    };

    db.servants.push(newServant);
    db.servants.sort((a, b) => {
      const codeA = parseInt(a.code) || 0;
      const codeB = parseInt(b.code) || 0;
      return codeA - codeB;
    });

    db.notifications.unshift({
      id: `notif-${Date.now()}`,
      title: "إضافة خادم عبر مزامنة Google Sheets",
      content: `تم إضافة الخادم ${newServant.name} (كود: ${newServant.code}) تلقائياً عبر جدول البيانات.`,
      type: "SUCCESS",
      date: `${new Date().toISOString().split("T")[0]} 12:00`,
      isRead: false
    });

    saveDB(db).catch(console.error);
    auditLog("system", "ADD", `إضافة خادم جديد تلقائياً عبر مزامنة الشيت: ${newServant.name} (كود: ${newServant.code})`);
    return res.json({ success: true, message: "تم إضافة الخادم الجديد بنجاح", data: newServant });
  } else {
    // Update existing servant
    const existing = db.servants[idx];

    existing.name = name ? String(name).trim() : existing.name;
    if (phone !== undefined) existing.phone = String(phone).trim();
    if (joinDate !== undefined) existing.joinDate = String(joinDate).trim();
    if (notes !== undefined) existing.notes = String(notes).trim();
    if (status !== undefined) existing.status = servantStatus;
    if (humanitarian !== undefined) {
      existing.humanitarian = isHumanitarian;
      if (humanitarianReason !== undefined) existing.humanitarianReason = String(humanitarianReason).trim();
    }
    if (isSupervisor !== undefined) existing.isSupervisor = isSuper;
    if (birthDate !== undefined) existing.birthDate = String(birthDate).trim();
    if (address !== undefined) existing.address = String(address).trim();

    db.notifications.unshift({
      id: `notif-${Date.now()}`,
      title: "تحديث خادم عبر مزامنة Google Sheets",
      content: `تم تحديث بيانات الخادم ${existing.name} (كود: ${existing.code}) بناءً على تعديله في جدول البيانات.`,
      type: "INFO",
      date: `${new Date().toISOString().split("T")[0]} 12:00`,
      isRead: false
    });

    saveDB(db).catch(console.error);
    auditLog("system", "EDIT", `تعديل خادم تلقائياً عبر مزامنة الشيت: ${existing.name} (كود: ${existing.code})`);
    return res.json({ success: true, message: "تم تحديث بيانات الخادم بنجاح", data: existing });
  }
});

// Endpoint triggered by Apps Script onChange for full synchronization and comparisons
app.post("/api/sheets-sync/trigger-pull", async (req, res) => {
  const db = loadDB();
  const activeUser = "Google Apps Script Trigger";

  try {
    const result = await syncGoogleSheetsData(undefined, activeUser);
    res.json(result);
  } catch (error: any) {
    console.error("Sheets Sync Trigger Pull Error:", error);
    res.status(500).json({
      success: false,
      message: error.message || "حدث خطأ أثناء المزامنة التلقائية للمسح والمقارنة"
    });
  }
});

// Helper to get Google Sheet's last modified time using Google Drive API
async function getSpreadsheetLastModifiedTime(spreadsheetId: string, accessToken?: string): Promise<Date | null> {
  return null;
}

function detectAndAccumulateChanges(db: any, clientDb: any) {
  if (!db.pendingChanges) db.pendingChanges = {};
  const p = db.pendingChanges;
  p.addedServants = p.addedServants || [];
  p.updatedServants = p.updatedServants || [];
  p.deletedServants = p.deletedServants || [];
  p.updatedDuties = p.updatedDuties || [];
  p.recordedAttendance = p.recordedAttendance || 0;
  p.addedMessages = p.addedMessages || 0;
  p.settingsUpdated = p.settingsUpdated || false;
  p.logoChanged = p.logoChanged || false;
  p.coverChanged = p.coverChanged || false;
  p.leavesUpdated = p.leavesUpdated || false;
  const now = new Date().toISOString();

  const touch = () => {
    if (!p.firstChangeTime) p.firstChangeTime = now;
    p.lastChangeTime = now;
  };

  // 1. Servants Comparison
  if (clientDb.servants && db.servants) {
    const existingMap = new Map<any, any>(db.servants.map((s: any) => [s.code, s]));
    const clientMap = new Map<any, any>(clientDb.servants.map((s: any) => [s.code, s]));

    for (const [code, cs] of clientMap.entries()) {
      const es = existingMap.get(code);
      if (!es) {
        if (!p.addedServants.includes(cs.name)) {
          p.addedServants.push(cs.name);
          touch();
        }
      } else {
        const isModified = 
          es.name !== cs.name ||
          es.phone !== cs.phone ||
          es.status !== cs.status ||
          es.isSupervisor !== cs.isSupervisor ||
          es.humanitarian !== cs.humanitarian ||
          es.humanitarianReason !== cs.humanitarianReason ||
          es.birthDate !== cs.birthDate ||
          es.address !== cs.address ||
          es.notes !== cs.notes ||
          es.avatar !== cs.avatar ||
          es.pinnedLocationId !== cs.pinnedLocationId;

        if (isModified) {
          if (!p.updatedServants.includes(cs.name)) {
            p.updatedServants.push(cs.name);
            touch();
          }
        }
      }
    }

    for (const [code, es] of existingMap.entries()) {
      if (!clientMap.has(code)) {
        if (!p.deletedServants.includes(es.name)) {
          p.deletedServants.push(es.name);
          touch();
        }
      }
    }
  }

  // 2. Settings Comparison
  if (clientDb.settings && db.settings) {
    const es = db.settings;
    const cs = clientDb.settings;

    if (cs.logoUrl && cs.logoUrl !== es.logoUrl) {
      if (!p.logoChanged) {
        p.logoChanged = true;
        touch();
      }
    }

    const esBanner = es.card?.bannerUrl;
    const csBanner = cs.card?.bannerUrl;
    if (csBanner && csBanner !== esBanner) {
      if (!p.coverChanged) {
        p.coverChanged = true;
        touch();
      }
    }

    const isSettingsModified = 
      cs.headerTitle !== es.headerTitle ||
      cs.subHeaderTitle !== es.subHeaderTitle ||
      cs.footerText !== es.footerText ||
      cs.allowSelfAttendance !== es.allowSelfAttendance ||
      cs.attendanceDeadline !== es.attendanceDeadline;

    if (isSettingsModified) {
      if (!p.settingsUpdated) {
        p.settingsUpdated = true;
        touch();
      }
    }
  }

  // 3. Duties Comparison
  if (clientDb.duties && db.duties) {
    const existingMap = new Map<any, any>(db.duties.map((d: any) => [d.id, d]));
    for (const cd of clientDb.duties) {
      const ed = existingMap.get(cd.id);
      if (!ed) {
        if (!p.updatedDuties.includes(cd.date)) {
          p.updatedDuties.push(cd.date);
          touch();
        }
      } else {
        const isModified = ed.date !== cd.date || ed.type !== cd.type || ed.notes !== cd.notes;
        if (isModified) {
          if (!p.updatedDuties.includes(cd.date)) {
            p.updatedDuties.push(cd.date);
            touch();
          }
        }
      }
    }
  }

  // 4. Attendance Comparison
  if (clientDb.attendance && db.attendance) {
    const existingCount = db.attendance.length;
    const clientCount = clientDb.attendance.length;
    if (clientCount > existingCount) {
      p.recordedAttendance += (clientCount - existingCount);
      touch();
    }
  }

  // 5. Messages Comparison
  if (clientDb.messages && db.messages) {
    const existingCount = db.messages.length;
    const clientCount = clientDb.messages.length;
    if (clientCount > existingCount) {
      p.addedMessages += (clientCount - existingCount);
      touch();
    }
  }

  // 6. Leaves Comparison
  if (clientDb.leaves && db.leaves) {
    const existingCount = db.leaves.length;
    const clientCount = clientDb.leaves.length;
    if (clientCount !== existingCount) {
      p.leavesUpdated = true;
      touch();
    }
  }
}

// -----------------------------------------------------------------
// MANUAL SYNC ENDPOINT
// -----------------------------------------------------------------
app.post("/api/sync/all", async (req, res) => {
  try {
    const clientDb = req.body;
    if (!clientDb) {
      return res.status(400).json({ success: false, message: "بيانات المزامنة مفقودة" });
    }
    
    const db = loadDB();
    
    // Accumulate changes before updating DB
    detectAndAccumulateChanges(db, clientDb);

    const keys = [
      "users", "servants", "duties", "locations", "distributions",
      "attendance", "leaves", "suspensions", "excluded", "messages",
      "notifications", "auditLogs", "sheetsSyncConfig", "sheetsSyncLogs",
      "financeCampaigns", "financeRecords", "honoringRecords", "settings"
    ];
    
    keys.forEach((key) => {
      if (clientDb[key] !== undefined) {
        if (key === "servants" && db.servants) {
          const serverServantsMap = new Map(db.servants.map((s: any) => [s.id, s]));
          clientDb.servants = clientDb.servants.map((cs: any) => {
             const ss = serverServantsMap.get(cs.id);
             // Protect existing base64 avatar from being overwritten by a stale client sync
             if (ss && ss.avatar && ss.avatar.startsWith("data:image")) {
                if (!cs.avatar || (!cs.avatar.startsWith("data:image") && cs.avatar !== ss.avatar)) {
                   console.log("[PHOTO_SAVED] db.json (Prevented stale client from overwriting base64 avatar for " + cs.name + ")");
                   return { ...cs, avatar: ss.avatar };
                }
             }
             return cs;
          });
        }
        (db as any)[key] = clientDb[key];
      }
    });
    
    saveDB(db).catch(console.error);
    
    // Automatic push is strictly prohibited. So we only save database locally.
    res.json({
      success: true,
      message: "تم حفظ التعديلات المحلية بنجاح في قاعدة البيانات المحلية. يرجى الضغط على زر (رفع البيانات) لترحيلها إلى Google Sheets.",
      db
    });
  } catch (error: any) {
    console.error("Manual sync all failed:", error);
    res.status(500).json({ success: false, message: error.message || "فشلت المزامنة" });
  }
});

// -----------------------------------------------------------------
// SYSTEM CONFIGURATION & BACKUP ENDPOINTS
// -----------------------------------------------------------------
app.get("/api/debug/storage-status", (req, res) => {
  const db = loadDB();
  const dbPath = path.join(process.cwd(), "db.json");
  
  let lastMod = "N/A";
  if (fs.existsSync(dbPath)) {
    lastMod = fs.statSync(dbPath).mtime.toISOString();
  }

  const envData = {
    runtime: process.env.NODE_ENV === "production" ? "Cloud Run / Shared" : "Preview",
    spreadsheetId: db.sheetsSyncConfig ? db.sheetsSyncConfig.spreadsheetId : "N/A",
    dbPath: dbPath,
    servantsCount: db.servants ? db.servants.length : 0,
    attendanceCount: db.attendance ? db.attendance.length : 0,
    lastModified: lastMod,
    sampleServant: db.servants && db.servants.length > 0 ? db.servants[0].name : "None"
  };

  console.log("[STORAGE_STATUS]");
  console.log("Runtime=" + envData.runtime);
  console.log("SpreadsheetId=" + envData.spreadsheetId);
  console.log("DB_PATH=" + envData.dbPath);
  console.log("ServantsCount=" + envData.servantsCount);
  console.log("AttendanceCount=" + envData.attendanceCount);

  res.json(envData);
});

app.get("/api/debug/environment", (req, res) => {
  const db = loadDB();
  const envData = {
    runtime: process.env.NODE_ENV === "production" ? "Cloud Run / Shared" : "Preview",
    spreadsheetId: db.sheetsSyncConfig ? db.sheetsSyncConfig.spreadsheetId : "N/A",
    dbPath: path.join(process.cwd(), "db.json"),
    projectId: process.env.GOOGLE_CLOUD_PROJECT || "N/A",
    timestamp: new Date().toISOString()
  };

  console.log("[ENVIRONMENT]");
  console.log("Runtime=" + envData.runtime);
  console.log("SpreadsheetId=" + envData.spreadsheetId);
  console.log("DB_PATH=" + envData.dbPath);
  console.log("ProjectId=" + envData.projectId);

  res.json(envData);
});

app.get("/api/settings", async (req, res) => {
  res.setHeader("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
  res.setHeader("Pragma", "no-cache");
  res.setHeader("Expires", "0");

  const db = loadDB();
  res.json(db.settings || {});
});

app.post("/api/settings", async (req, res) => {
  const db = loadDB();
  db.settings = {
    ...db.settings,
    ...req.body
  };
  saveDB(db).catch(console.error);
  auditLog("admin", "EDIT", "تحديث إعدادات الصفحة وسمة النظام وصورة الغلاف");

  // Push updated settings back to Google Sheets
  pushSettingsToGoogleSheet(db.settings, true).catch(e => {
    if (e.message && e.message.includes("لم يتم العثور على صلاحيات")) {
      console.log("[Push Settings] Ignored: No Google Sheets credentials configured.");
    } else {
      console.warn("[Push Settings Warning during save]", e.message || e);
    }
  });

  res.json({ success: true, settings: db.settings });
});

app.get("/api/backup/download", async (req, res) => {
  const db = loadDB();
  res.setHeader("Content-Type", "application/json");
  res.setHeader("Content-Disposition", "attachment; filename=karbala-system-backup.json");
  res.send(JSON.stringify(db, null, 2));
});

app.post("/api/backup/restore", async (req, res) => {
  const backupData = req.body;
  
  if (!backupData || !backupData.servants || !backupData.users) {
    return res.status(400).json({
      success: false,
      message: "ملف النسخة الاحتياطية غير صالح أو ناقص البيانات الأساسية."
    });
  }
  
  const db = loadDB();
  
  // Merge or completely overwrite? Overwriting is cleaner, but let's make sure we preserve the existing administrator.
  const originalAdmin = db.users.find(u => u.role === "ADMIN");
  
  db.users = backupData.users || db.users;
  if (originalAdmin && !db.users.some(u => u.role === "ADMIN")) {
    db.users.push(originalAdmin);
  }
  
  db.servants = backupData.servants || db.servants;
  db.duties = backupData.duties || db.duties;
  db.locations = backupData.locations || db.locations;
  db.distributions = backupData.distributions || db.distributions;
  db.attendance = backupData.attendance || db.attendance;
  db.leaves = backupData.leaves || db.leaves;
  db.suspensions = backupData.suspensions || db.suspensions;
  db.excluded = backupData.excluded || db.excluded;
  db.messages = backupData.messages || db.messages;
  db.notifications = backupData.notifications || db.notifications;
  db.auditLogs = backupData.auditLogs || db.auditLogs;
  db.sheetsSyncConfig = backupData.sheetsSyncConfig || db.sheetsSyncConfig;
  db.sheetsSyncLogs = backupData.sheetsSyncLogs || db.sheetsSyncLogs;
  db.financeCampaigns = backupData.financeCampaigns || db.financeCampaigns;
  db.financeRecords = backupData.financeRecords || db.financeRecords;
  db.settings = backupData.settings || db.settings;
  
  saveDB(db).catch(console.error);
  auditLog("admin", "EDIT", "استعادة وتطبيق نسخة احتياطية كاملة للنظام");
  res.json({ success: true });
});

// -----------------------------------------------------------------
// VITE DEV SERVER & PRODUCTION ROUTING
// -----------------------------------------------------------------
let isPushingToSheets = false; (global as any).isPushingToSheets = false;
let pendingSheetsPush = false;
let latestDbForPush: any = null;



async function startServer() {
  console.log("-----------------------------------------");
  console.log("Starting Server...");
  await initFirestoreDB();
  const tmpDb = loadDB();
  console.log("[ENVIRONMENT]");
  console.log("ProjectId=" + (process.env.GOOGLE_CLOUD_PROJECT || "N/A"));
  console.log("SpreadsheetId=" + (tmpDb.sheetsSyncConfig ? tmpDb.sheetsSyncConfig.spreadsheetId : "N/A"));
  console.log("DB_PATH=" + path.join(process.cwd(), "db.json"));
  console.log("Runtime=" + (process.env.NODE_ENV === "production" ? "Cloud Run / Shared" : "Preview"));
  const DATA_SOURCE = "google-sheets";
  console.log("Current data source:", DATA_SOURCE);
  
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", async (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://0.0.0.0:${PORT}`);
    
    // Background task: Initialize data from Google Sheets after server binds to port
    
    console.log("Checking Data Source...");
    dbReadyPromise = (async () => {

      try {
         if (process.env.ENABLE_GOOGLE_SHEETS_SYNC === "true") {
           console.log("[Startup] Syncing state from Google Sheets as Primary Source of Truth...");
           await syncGoogleSheetsData();
           (global as any).isSyncingFromSheets = false; const nowLog = new Date().toISOString();
           console.log(`[${nowLog}] [READ] syncGoogleSheetsData -> Google Sheets`);
         } else {
           console.log("[Startup] ENABLE_GOOGLE_SHEETS_SYNC is not 'true'. Skipping startup Google Sheets sync.");
         }
      } catch (err) {
         console.warn("⚠️ [Startup] Failed to read from Google Sheets. Falling back to local db.json temporary cache.", String(err));
         console.warn("Startup source: db.json");
      }

      // Run background scan on system startup to catch updates
      runSystemScans(loadDB());
      isDbReady = true;
    })();

  });
}


// -----------------------------------------------------------------
// AUTO SAVE TO GOOGLE SHEETS
// -----------------------------------------------------------------
let isSyncingToSheets = false;

setInterval(async () => {
    if (process.env.ENABLE_GOOGLE_SHEETS_SYNC !== "true") {
        return; // Skip automatic sync interval entirely
    }
    if (!hasPendingChanges) return;

    if (isSyncingToSheets) return;

    try {
        isSyncingToSheets = true;

        console.log("[AUTO SAVE] Pending changes found");
        console.log("[AUTO SAVE] Uploading to Google Sheets");

        await pushAllServantsToGoogleSheet(true);

        setHasPendingChanges(false);

        console.log("[AUTO SAVE] Upload completed");
    } catch (error) {
        console.error("[AUTO SAVE ERROR]", error);
    } finally {
        isSyncingToSheets = false;
    }
}, 600000);

startServer();
