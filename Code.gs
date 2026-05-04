/**
 * ==========================================
 * TRIP PLANNER SUPERAPP BACKEND
 * By: Iwan Tingwe
 * ==========================================
 */

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Trip App - SuperApp & Desktop')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

// ==========================================
// SECURITY & HASHING (SHA-256 with Salt)
// ==========================================

/**
 * Fungsi untuk melakukan hashing password dengan salt
 */
function hashPassword(password, salt) {
  const combined = password + salt;
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, combined);
  let hash = '';
  for (let i = 0; i < digest.length; i++) {
    let byte = digest[i];
    if (byte < 0) byte += 256;
    let hex = byte.toString(16);
    if (hex.length === 1) hex = '0' + hex;
    hash += hex;
  }
  return hash;
}

/**
 * JALANKAN FUNGSI INI 1X UNTUK MEMBUAT DATABASE DAN MENGISI DATA DUMMY
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const schemas = {
    "Users": ["email", "password", "salt", "name", "created_at"],
    "Trips": ["trip_id", "trip_name", "description", "start_date", "end_date", "budget", "currency", "creator", "image", "status"],
    "Destinations": ["dest_id", "trip_id", "date", "time", "place_name", "notes", "created_by", "calendar_event_id"],
    "Expenses": ["expense_id", "trip_id", "description", "amount", "currency", "paid_by", "date", "category", "proof_image"],
    "Members": ["member_id", "trip_id", "email", "name", "role", "status"],
    "Alerts": ["alert_id", "email", "title", "message", "type", "date", "is_read", "trip_id"],
    "Tasks": ["task_id", "trip_id", "title", "category", "due_date", "assigned_to", "completed", "created_by", "created_at"],
    "AppSettings": ["key", "value", "description"]
  };

  // Membuat sheet & kolom jika belum ada, dan PAKSA update header agar selalu sinkron
  for (let sheetName in schemas) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    // Paksa update baris 1 (Header) dengan schema baku untuk cegah mismatch kolom
    sheet.getRange(1, 1, 1, schemas[sheetName].length)
         .setValues([schemas[sheetName]])
         .setFontWeight("bold")
         .setBackground("#00D1FF");
  }

  // --- PROSES PEMBERSIHAN & PENGISIAN DATA DUMMY ---
  const now = new Date();
  
  // 1. Membersihkan & Mengisi Users
  const userSheet = ss.getSheetByName("Users");
  if (userSheet.getLastRow() > 1) userSheet.deleteRows(2, userSheet.getLastRow() - 1);
  const users = [
    { email: "admin@trip.com", pass: "admin123", name: "Iwan Tingwe (Admin)" },
    { email: "manager@travel.id", pass: "manage789", name: "Budi Santoso" },
    { email: "traveler@gmail.com", pass: "travel456", name: "Siti Aminah" }
  ];
  users.forEach(u => {
    const salt = Utilities.getUuid().split('-')[0];
    userSheet.appendRow([u.email, hashPassword(u.pass, salt), salt, u.name, now]);
  });

  // 2. Membersihkan & Mengisi Trips
  const tripSheet = ss.getSheetByName("Trips");
  if (tripSheet.getLastRow() > 1) tripSheet.deleteRows(2, tripSheet.getLastRow() - 1);
  const tripId = "t_jepang_2026";
  tripSheet.appendRow([
    tripId, 
    "Liburan Musim Semi di Jepang", 
    "Menikmati bunga Sakura di Tokyo dan Kyoto bersama tim.", 
    "2026-04-10", 
    "2026-04-15", 
    50000000, 
    "Rp", 
    "admin@trip.com", 
    "https://images.unsplash.com/photo-1493976040374-85c8e12f0c0e?q=80&w=1200&auto=format&fit=crop"
  ]);

  // 3. Membersihkan & Mengisi Destinations (Timeline)
  const destSheet = ss.getSheetByName("Destinations");
  if (destSheet.getLastRow() > 1) destSheet.deleteRows(2, destSheet.getLastRow() - 1);
  const dests = [
    [tripId + "_d1", tripId, "2026-04-10", "10:00", "Landing at Narita Airport", "Ambil JR Pass di loket", "admin@trip.com", ""],
    [tripId + "_d2", tripId, "2026-04-10", "14:00", "Check-in Hotel Shinjuku", "Istirahat sejenak", "admin@trip.com", ""],
    [tripId + "_d3", tripId, "2026-04-10", "19:00", "Dinner at Omoide Yokocho", "Coba Yakitori khas Tokyo", "admin@trip.com", ""],
    [tripId + "_d4", tripId, "2026-04-11", "09:00", "Medical Checkup Tokyo Clinic", "Sesuai janji temu medis", "admin@trip.com", ""],
    [tripId + "_d5", tripId, "2026-04-11", "13:00", "Shinjuku Gyoen National Garden", "Melihat Sakura mekar", "admin@trip.com", ""]
  ];
  dests.forEach(d => destSheet.appendRow(d));

  // 4. Membersihkan & Mengisi Expenses
  const expSheet = ss.getSheetByName("Expenses");
  if (expSheet.getLastRow() > 1) expSheet.deleteRows(2, expSheet.getLastRow() - 1);
  const exps = [
    [tripId + "_e1", tripId, "Tiket Pesawat Garuda PP", 12000000, "Rp", "admin@trip.com", "2026-03-01", "Transportasi"],
    [tripId + "_e2", tripId, "Booking Airbnb 5 Malam", 15000000, "Rp", "manager@travel.id", "2026-03-05", "Akomodasi"],
    [tripId + "_e3", tripId, "Makan Siang Ramen", 500000, "Rp", "traveler@gmail.com", "2026-04-10", "Makanan"]
  ];
  exps.forEach(e => expSheet.appendRow(e));

  // 5. Membersihkan & Mengisi Members
  const memSheet = ss.getSheetByName("Members");
  if (memSheet.getLastRow() > 1) memSheet.deleteRows(2, memSheet.getLastRow() - 1);
  const mems = [
    [tripId + "_m1", tripId, "admin@trip.com", "Iwan Tingwe (Admin)", "owner", "active"],
    [tripId + "_m2", tripId, "manager@travel.id", "Budi Santoso", "admin", "active"],
    [tripId + "_m3", tripId, "traveler@gmail.com", "Siti Aminah", "member", "active"]
  ];
  mems.forEach(m => memSheet.appendRow(m));

  // 6. Membersihkan & Mengisi Alerts
  const alertSheet = ss.getSheetByName("Alerts");
  if (alertSheet.getLastRow() > 1) alertSheet.deleteRows(2, alertSheet.getLastRow() - 1);
  const alerts = [
    ["a1", "admin@trip.com", "Selamat Datang!", "Mulai rencanakan trip impianmu.", "info", now.toISOString(), "false"],
    ["a2", "admin@trip.com", "Pengeluaran Baru", "Budi Santoso menambahkan biaya Akomodasi.", "expense", now.toISOString(), "false"]
  ];
  alerts.forEach(a => alertSheet.appendRow(a));

  return "Database Berhasil Direset! Semua data dummy telah disinkronkan.";
}


// --- HELPER UNTUK MENGAMBIL DATA DALAM BENTUK OBJECT ---
function getSheetDataAsObjects(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const result = [];
  for (let i = 1; i < data.length; i++) {
    let obj = {};
    for (let j = 0; j < headers.length; j++) {
      let value = data[i][j];
      
      // Mencegah error serialisasi dari Google Apps Script ke Frontend
      if (value instanceof Date) {
        if (headers[j] === "time") {
          let h = ("0" + value.getHours()).slice(-2);
          let m = ("0" + value.getMinutes()).slice(-2);
          value = `${h}:${m}`;
        } else {
          let y = value.getFullYear();
          let m = ("0" + (value.getMonth() + 1)).slice(-2);
          let d = ("0" + value.getDate()).slice(-2);
          value = `${y}-${m}-${d}`;
        }
      } else if (value === null || value === undefined) {
        value = ""; // Pastikan sel kosong menjadi string kosong
      } else if (typeof value !== 'string' && typeof value !== 'number' && typeof value !== 'boolean') {
        // Tangkap objek #N/A (Error cell) agar tidak merusak proses fetching JSON
        value = String(value);
      }
      
      obj[headers[j]] = value;
    }
    obj._rowIndex = i + 1;
    result.push(obj);
  }
  return result;
}


// ==========================================
// API / CRUD FUNCTIONS
// ==========================================

// ==========================================
// SESSION MANAGEMENT (CacheService)
// ==========================================
const SESSION_DURATION = 1800; // 30 Menit dalam detik

/**
 * Fungsi untuk membuat session baru
 */
function createSession(userData) {
  const token = "sess_" + Utilities.getUuid();
  const cache = CacheService.getUserCache();
  
  // Simpan data user ke cache
  cache.put(token, JSON.stringify(userData), SESSION_DURATION);
  return token;
}

/**
 * Fungsi untuk mengecek apakah session masih valid
 */
function checkSession(token) {
  if (!token) return { success: false };
  
  const cache = CacheService.getUserCache();
  const cachedData = cache.get(token);
  
  if (cachedData) {
    // Perpanjang session setiap kali ada aktivitas (optional)
    const userData = JSON.parse(cachedData);
    cache.put(token, cachedData, SESSION_DURATION);
    return { success: true, data: userData };
  }
  
  return { success: false, message: "Session telah berakhir" };
}

/**
 * Fungsi untuk keluar dari aplikasi (Hapus Session)
 */
function logout(token) {
  if (token) {
    const cache = CacheService.getUserCache();
    cache.remove(token);
  }
  return { success: true, message: "Berhasil keluar" };
}

function login(data) {
  try {
    const users = getSheetDataAsObjects("Users");
    const email = String(data.email).trim().toLowerCase();
    const foundUser = users.find(u => String(u.email).trim().toLowerCase() === email);
    
    if (foundUser) {
      // Verifikasi Password dengan Hash & Salt
      const hashedInput = hashPassword(String(data.password), foundUser.salt);
      
      if (hashedInput === String(foundUser.password)) {
        const userData = { email: foundUser.email, name: foundUser.name };
        const token = createSession(userData);
        return { 
          success: true, 
          data: userData, 
          token: token,
          message: "Berhasil masuk!" 
        };
      }
    }
    return { success: false, message: "Email atau Password salah!" };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getCurrentUser(token) {
  const session = checkSession(token);
  return session.success ? session.data : null; 
}

function getMyTrips(userEmail) {
  try {
    const trips = getSheetDataAsObjects("Trips");
    const members = getSheetDataAsObjects("Members");
    
    // Ambil tripId dimana user tersebut adalah member AKTIF
    const myTripIds = members
      .filter(m => String(m.email).trim().toLowerCase() === String(userEmail).trim().toLowerCase() && m.status === 'active')
      .map(m => String(m.trip_id));
    
    // Filter trips: trips yang dibuat user ATAU user adalah member aktif
    const filteredTrips = trips.filter(t => String(t.creator).trim().toLowerCase() === String(userEmail).trim().toLowerCase() || myTripIds.includes(String(t.trip_id)));
    
    return { success: true, data: filteredTrips };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function createTrip(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tripId = "t" + Date.now();
    const imgUrl = "https://images.unsplash.com/photo-1469854523086-cc02fe5d8800?q=80&w=600&auto=format&fit=crop";
    
    ss.getSheetByName("Trips").appendRow([
      tripId, data.trip_name || "", data.description || "", data.start_date || "", data.end_date || "", data.budget || 0, data.currency || "Rp", data.creator || "", imgUrl, "active"
    ]);
    
    // Otomatis jadikan creator sebagai Owner
    ss.getSheetByName("Members").appendRow([
      "m" + Date.now(), tripId, data.creator || "", data.creatorName || "", "owner", "active"
    ]);
    
    return { success: true, tripId: tripId, message: "Trip Berhasil Dibuat" };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function updateTrip(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Trips");
    const trips = getSheetDataAsObjects("Trips");
    const target = trips.find(t => String(t.trip_id) === String(data.trip_id));
    
    if (target) {
      sheet.getRange(target._rowIndex, 2).setValue(data.trip_name || "");
      sheet.getRange(target._rowIndex, 3).setValue(data.description || "");
      sheet.getRange(target._rowIndex, 4).setValue(data.start_date || "");
      sheet.getRange(target._rowIndex, 5).setValue(data.end_date || "");
      sheet.getRange(target._rowIndex, 6).setValue(data.budget || 0);
      sheet.getRange(target._rowIndex, 7).setValue(data.currency || "Rp");
      if (data.status) sheet.getRange(target._rowIndex, 10).setValue(data.status);
      return { success: true, message: "Trip diperbarui" };
    }
    return { success: false, message: "Trip tidak ditemukan" };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function deleteTrip(tripId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Trips");
    const trips = getSheetDataAsObjects("Trips");
    const target = trips.find(t => String(t.trip_id) === String(tripId));
    
    if (target) {
      sheet.deleteRow(target._rowIndex);
      return { success: true, message: "Trip dihapus" };
    }
    return { success: false, message: "Gagal menghapus" };
  } catch(e) {
    return { success: false, message: e.message };
  }
}


// --- ITINERARY (DESTINATIONS) ---
function getItinerary(tripId) {
  try {
    const dests = getSheetDataAsObjects("Destinations").filter(d => String(d.trip_id) === String(tripId));
    dests.sort((a, b) => String(a.date).localeCompare(String(b.date)) || String(a.time).localeCompare(String(b.time)));
    return { success: true, data: dests };
  } catch(e) { return { success: false, message: e.message }; }
}

function addItinerary(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Destinations");
    sheet.appendRow([
      "d" + Date.now(), data.trip_id || "", data.date || "", data.time || "", data.place_name || "", data.notes || "", data.created_by || "", ""
    ]);
    return { success: true, message: "Agenda ditambahkan" };
  } catch(e) { return { success: false, message: e.message }; }
}

function updateItinerary(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Destinations");
    const dests = getSheetDataAsObjects("Destinations");
    const target = dests.find(d => String(d.dest_id) === String(data.dest_id));
    if (target) {
      sheet.getRange(target._rowIndex, 3, 1, 4).setValues([[data.date || "", data.time || "", data.place_name || "", data.notes || ""]]);
      return { success: true, message: "Agenda diperbarui" };
    }
    return { success: false, message: "Gagal edit agenda" };
  } catch(e) { return { success: false, message: e.message }; }
}

function deleteItinerary(id) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Destinations");
    const target = getSheetDataAsObjects("Destinations").find(d => String(d.dest_id) === String(id));
    if (target) {
      sheet.deleteRow(target._rowIndex);
      return { success: true, message: "Agenda dihapus" };
    }
    return { success: false, message: "Gagal hapus" };
  } catch(e) { return { success: false, message: e.message }; }
}

function syncToCalendar(id) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Destinations");
    const target = getSheetDataAsObjects("Destinations").find(d => String(d.dest_id) === String(id));
    if (target) {
      sheet.getRange(target._rowIndex, 8).setValue("sync_" + Date.now());
      return { success: true, message: "Disinkronkan dengan Calendar!" };
    }
    return { success: false, message: "Gagal sinkronisasi" };
  } catch(e) { return { success: false, message: e.message }; }
}


// --- PENGELUARAN (EXPENSES) ---
function getExpenses(tripId) {
  try {
    const exps = getSheetDataAsObjects("Expenses").filter(e => String(e.trip_id) === String(tripId));
    exps.sort((a, b) => String(b.date).localeCompare(String(a.date)));
    return { success: true, data: exps };
  } catch(e) { return { success: false, message: e.message }; }
}

/**
 * Fungsi untuk upload file ke Google Drive dan mengembalikan link view
 */
function uploadToDrive(base64Data, fileName) {
  try {
    const folderName = "TripPlanner_Proofs";
    let folders = DriveApp.getFoldersByName(folderName);
    let folder;
    
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
    }
    
    const contentType = base64Data.substring(base64Data.indexOf(":") + 1, base64Data.indexOf(";"));
    const bytes = Utilities.base64Decode(base64Data.split(",")[1]);
    const blob = Utilities.newBlob(bytes, contentType, fileName);
    
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return file.getUrl();
  } catch (e) {
    throw new Error("Gagal upload ke Drive: " + e.message);
  }
}

function addExpense(data) {
  try {
    let proofUrl = data.proof_image || "";
    
    // Jika ada file bukti yang diupload (base64)
    if (data.fileData && data.fileName) {
      proofUrl = uploadToDrive(data.fileData, data.fileName);
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Expenses");
    sheet.appendRow([
      "e" + Date.now(), data.trip_id || "", data.description || "", data.amount || 0, data.currency || "Rp", data.paid_by || "", data.date || "", data.category || "", proofUrl
    ]);
    return { success: true, message: "Transaksi dicatat" };
  } catch(e) { return { success: false, message: e.message }; }
}

function updateExpense(data) {
  try {
    let proofUrl = data.proof_image || "";
    
    // Jika ada file bukti baru yang diupload
    if (data.fileData && data.fileName) {
      proofUrl = uploadToDrive(data.fileData, data.fileName);
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Expenses");
    const exps = getSheetDataAsObjects("Expenses");
    const target = exps.find(e => String(e.expense_id) === String(data.expense_id));
    if (target) {
      sheet.getRange(target._rowIndex, 3, 1, 7).setValues([[data.description || "", data.amount || 0, data.currency || "Rp", data.paid_by || "", data.date || "", data.category || "", proofUrl]]);
      return { success: true, message: "Transaksi diperbarui" };
    }
    return { success: false, message: "Gagal edit" };
  } catch(e) { return { success: false, message: e.message }; }
}

function deleteExpense(id) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Expenses");
    const target = getSheetDataAsObjects("Expenses").find(e => String(e.expense_id) === String(id));
    if (target) {
      sheet.deleteRow(target._rowIndex);
      return { success: true, message: "Transaksi dihapus" };
    }
    return { success: false, message: "Gagal hapus" };
  } catch(e) { return { success: false, message: e.message }; }
}


// --- TASKS / CHECKLIST ---
function getTasks(tripId) {
  try {
    const tasks = getSheetDataAsObjects("Tasks").filter(t => String(t.trip_id) === String(tripId));
    tasks.sort((a, b) => new Date(b.created_at) - new Date(a.created_at));
    return { success: true, data: tasks };
  } catch(e) { return { success: false, message: e.message }; }
}

function addTask(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tasks");
    sheet.appendRow([
      "task_" + Date.now(),
      data.trip_id,
      data.title || "",
      data.category || "Umum",
      data.due_date || "",
      data.assigned_to || "",
      false, // completed
      data.created_by || "",
      new Date().toISOString()
    ]);
    return { success: true, message: "Tugas ditambahkan" };
  } catch(e) { return { success: false, message: e.message }; }
}

function updateTask(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tasks");
    const tasks = getSheetDataAsObjects("Tasks");
    const target = tasks.find(t => String(t.task_id) === String(data.task_id));
    if (target) {
      sheet.getRange(target._rowIndex, 3, 1, 4).setValues([[
        data.title || "",
        data.category || "Umum",
        data.due_date || "",
        data.assigned_to || ""
      ]]);
      return { success: true, message: "Tugas diperbarui" };
    }
    return { success: false, message: "Tugas tidak ditemukan" };
  } catch(e) { return { success: false, message: e.message }; }
}

function toggleTaskStatus(taskId, completed) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tasks");
    const tasks = getSheetDataAsObjects("Tasks");
    const target = tasks.find(t => String(t.task_id) === String(taskId));
    if (target) {
      sheet.getRange(target._rowIndex, 7).setValue(completed);
      return { success: true, message: completed ? "Tugas selesai" : "Tugas dibuka kembali" };
    }
    return { success: false, message: "Tugas tidak ditemukan" };
  } catch(e) { return { success: false, message: e.message }; }
}

function deleteTask(taskId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tasks");
    const tasks = getSheetDataAsObjects("Tasks");
    const target = tasks.find(t => String(t.task_id) === String(taskId));
    if (target) {
      sheet.deleteRow(target._rowIndex);
      return { success: true, message: "Tugas dihapus" };
    }
    return { success: false, message: "Gagal menghapus" };
  } catch(e) { return { success: false, message: e.message }; }
}

// --- BALANCES / SETTLE UP LOGIC ---
function getBalances(tripId) {
  try {
    const expenses = getSheetDataAsObjects("Expenses").filter(e => String(e.trip_id) === String(tripId));
    const members = getSheetDataAsObjects("Members").filter(m => String(m.trip_id) === String(tripId));
    const trip = getSheetDataAsObjects("Trips").find(t => String(t.trip_id) === String(tripId));
    
    let totalExpense = 0;
    let totalIncome = 0; // Total dari Settle Up
    let b = {}; 
    
    const memCount = members.length || 1; 
    const perPersonBudget = trip ? (Number(trip.budget) || 0) : 0; // Total budget adalah target per orang
    
    // Inisialisasi: Semua member mulai dengan hutang sebesar target budget per orang
    members.forEach(m => b[m.email] = -perPersonBudget); 
    
    expenses.forEach(e => { 
      let amt = Number(e.amount) || 0; 
      if (e.category === 'Settle Up') {
         totalIncome += amt;
         if(b[e.paid_by] !== undefined) b[e.paid_by] += amt;
         // Settle Up mengurangi hutang user ke grup (biasanya disetor ke creator/kas)
         if(trip && b[trip.creator] !== undefined && e.paid_by !== trip.creator) {
            // b[trip.creator] -= amt; // Opsional: jika ingin melacak kas di creator
         }
      } else {
         totalExpense += amt; 
         if(b[e.paid_by] !== undefined) b[e.paid_by] += amt; 
         // Pengeluaran dibagi rata, mengurangi saldo semua orang
         Object.keys(b).forEach(k => b[k] -= (amt / memCount)); 
      }
    }); 
    
    // Perbaikan: Saldo akhir disesuaikan agar jika total pengeluaran masih di bawah budget, 
    // sisa budget tidak dianggap sebagai hutang tambahan yang harus dibayar sekarang.
    // Namun user minta "awal semua setting tidak lunas", jadi inisialisasi di atas sudah benar.
    
    let res = Object.keys(b).map(k => { 
      const mem = members.find(m => m.email === k); 
      return { email: k, name: mem ? mem.name : k.split("@")[0], balance: b[k] }; 
    }); 
    
    return { 
      success: true, 
      data: res, 
      total_expense: totalExpense, 
      total_income: totalIncome,
      total_budget: trip ? (Number(trip.budget) || 0) : 0,
      per_person_budget: perPersonBudget,
      total_members: memCount 
    };
  } catch(e) {
    return { success: false, message: e.message };
  }
}


function acceptInvitation(alertId, tripId, email) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const memSheet = ss.getSheetByName("Members");
    const alerts = getSheetDataAsObjects("Alerts");
    const members = getSheetDataAsObjects("Members");
    
    // 1. Cari record member yang pending
    const targetMem = members.find(m => String(m.trip_id) === String(tripId) && String(m.email).toLowerCase() === String(email).toLowerCase());
    if (targetMem) {
      memSheet.getRange(targetMem._rowIndex, 6).setValue("active");
    }
    
    // 2. Tandai alert sebagai terbaca
    const targetAlert = alerts.find(a => String(a.alert_id) === String(alertId));
    if (targetAlert) {
      ss.getSheetByName("Alerts").getRange(targetAlert._rowIndex, 7).setValue("true");
    }
    
    return { success: true, message: "Undangan diterima!" };
  } catch(e) { return { success: false, message: e.message }; }
}

function declineInvitation(alertId, tripId, email) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const memSheet = ss.getSheetByName("Members");
    const alerts = getSheetDataAsObjects("Alerts");
    const members = getSheetDataAsObjects("Members");
    
    // 1. Hapus record member
    const targetMem = members.find(m => String(m.trip_id) === String(tripId) && String(m.email).toLowerCase() === String(email).toLowerCase());
    if (targetMem) {
      memSheet.deleteRow(targetMem._rowIndex);
    }
    
    // 2. Tandai alert sebagai terbaca
    const targetAlert = alerts.find(a => String(a.alert_id) === String(alertId));
    if (targetAlert) {
      ss.getSheetByName("Alerts").getRange(targetAlert._rowIndex, 7).setValue("true");
    }
    
    return { success: true, message: "Undangan ditolak" };
  } catch(e) { return { success: false, message: e.message }; }
}

function changePassword(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName("Users");
    const users = getSheetDataAsObjects("Users");
    const email = String(data.email).trim().toLowerCase();
    const foundUser = users.find(u => String(u.email).trim().toLowerCase() === email);
    
    if (!foundUser) return { success: false, message: "User tidak ditemukan" };
    
    // 1. Verifikasi Password Lama
    const oldHashed = hashPassword(String(data.oldPassword), foundUser.salt);
    if (oldHashed !== String(foundUser.password)) {
      return { success: false, message: "Password lama salah!" };
    }
    
    // 2. Hash Password Baru (Gunakan salt yang sama atau baru, di sini kita gunakan salt yang sudah ada)
    const newHashed = hashPassword(String(data.newPassword), foundUser.salt);
    
    // 3. Update di Sheet
    userSheet.getRange(foundUser._rowIndex, 2).setValue(newHashed);
    
    return { success: true, message: "Password berhasil diperbarui!" };
  } catch(e) { return { success: false, message: e.message }; }
}

// --- GOOGLE CALENDAR SYNC ---
function syncTripToCalendar(tripId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const trips = getSheetDataAsObjects("Trips");
    const trip = trips.find(t => String(t.trip_id) === String(tripId));
    if (!trip) return { success: false, message: "Trip tidak ditemukan" };

    const destinations = getSheetDataAsObjects("Destinations").filter(d => String(d.trip_id) === String(tripId));
    if (destinations.length === 0) return { success: false, message: "Tidak ada jadwal (Timeline) untuk disinkronisasi." };

    const calendar = CalendarApp.getDefaultCalendar();
    let count = 0;

    destinations.forEach(dest => {
      const eventTitle = `[${trip.trip_name}] ${dest.place_name}`;
      const eventDate = new Date(dest.date);
      const timeParts = (dest.time || "09:00").split(":");
      
      const startTime = new Date(eventDate.getFullYear(), eventDate.getMonth(), eventDate.getDate(), parseInt(timeParts[0]), parseInt(timeParts[1]));
      const endTime = new Date(startTime.getTime() + (60 * 60 * 1000)); // Default 1 jam

      // Cek apakah event sudah ada (opsional, untuk mencegah duplikat sederhana)
      const existingEvents = calendar.getEvents(startTime, endTime, {search: eventTitle});
      if (existingEvents.length === 0) {
        calendar.createEvent(eventTitle, startTime, endTime, {
          description: dest.notes || "",
          location: dest.place_name
        });
        count++;
      }
    });

    return { success: true, message: `${count} jadwal berhasil disinkronkan ke Google Calendar Anda!` };
  } catch(e) { 
    if (e.message.indexOf("Authorization") > -1 || e.message.indexOf("permission") > -1) {
      return { success: false, message: "Otorisasi diperlukan. Silakan jalankan fungsi syncTripToCalendar sekali di Editor GAS untuk memberikan izin." };
    }
    return { success: false, message: "Gagal Sinkronisasi: " + e.message }; 
  }
}

// --- FULL DATA (Optimization) ---
function getTripFullData(tripId) {
  try {
    const itinerary = getItinerary(tripId);
    const expenses = getExpenses(tripId);
    const balances = getBalances(tripId);
    const members = getMembers(tripId);
    const tasks = getTasks(tripId);
    
    return {
      success: true,
      itinerary: itinerary.success ? itinerary.data : [],
      expenses: expenses.success ? expenses.data : [],
      balances: balances.success ? balances : { data: [], total_expense: 0, total_income: 0, total_budget: 0, per_person_budget: 0 },
      members: members.success ? members.data : [],
      tasks: tasks.success ? tasks.data : []
    };
  } catch(e) { return { success: false, message: e.message }; }
}

function getAppSettings() {
  try {
    const settings = getSheetDataAsObjects("AppSettings");
    const result = {};
    settings.forEach(s => { result[s.key] = s.value; });
    
    // Default values jika kosong
    if (!result.app_name) result.app_name = "TripSync";
    if (!result.primary_color) result.primary_color = "#00D1FF";
    
    return { success: true, data: result };
  } catch(e) { return { success: false, message: e.message }; }
}

function updateAppSettings(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("AppSettings");
    const settings = getSheetDataAsObjects("AppSettings");
    
    Object.keys(data).forEach(key => {
      const target = settings.find(s => s.key === key);
      if (target) {
        sheet.getRange(target._rowIndex, 2).setValue(data[key]);
      } else {
        sheet.appendRow([key, data[key], ""]);
      }
    });
    
    return { success: true, message: "Pengaturan aplikasi diperbarui" };
  } catch(e) { return { success: false, message: e.message }; }
}

// --- MEMBERS ---
function getMembers(tripId) {
  try {
    const mems = getSheetDataAsObjects("Members").filter(m => String(m.trip_id) === String(tripId));
    return { success: true, data: mems };
  } catch(e) { return { success: false, message: e.message }; }
}

// --- ALERTS / NOTIFICATIONS ---
function getMyAlerts(email) {
  try {
    const alerts = getSheetDataAsObjects("Alerts").filter(a => String(a.email).trim().toLowerCase() === String(email).trim().toLowerCase());
    alerts.sort((a, b) => new Date(b.date) - new Date(a.date));
    return { success: true, data: alerts };
  } catch(e) { return { success: false, message: e.message }; }
}

function createAlert(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Alerts");
    sheet.appendRow([
      "a" + Date.now(), data.email || "", data.title || "", data.message || "", data.type || "info", new Date().toISOString(), "false", data.trip_id || ""
    ]);
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

function markAlertAsRead(alertId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Alerts");
    const alerts = getSheetDataAsObjects("Alerts");
    const target = alerts.find(a => String(a.alert_id) === String(alertId));
    if (target) {
      sheet.getRange(target._rowIndex, 7).setValue("true");
      return { success: true };
    }
    return { success: false };
  } catch(e) { return { success: false }; }
}

function addMember(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const memSheet = ss.getSheetByName("Members");
    const userSheet = ss.getSheetByName("Users");
    
    const mems = getSheetDataAsObjects("Members").filter(m => String(m.trip_id) === String(data.trip_id));
    if (mems.some(m => String(m.email).toLowerCase() === String(data.email).toLowerCase())) {
      return { success: false, message: "Email sudah terdaftar di Trip ini!" };
    }
    
    const email = String(data.email).trim().toLowerCase();
    const users = getSheetDataAsObjects("Users");
    const existingUser = users.find(u => String(u.email).trim().toLowerCase() === email);
    
    let tempPassword = "";
    let isNewUser = false;
    let namaUser = data.email.split("@")[0];

    if (!existingUser) {
      isNewUser = true;
      // Buat Password Temporary (8 karakter random)
      tempPassword = Math.random().toString(36).slice(-8);
      const salt = Utilities.getUuid().split('-')[0];
      const hashedPass = hashPassword(tempPassword, salt);
      
      userSheet.appendRow([
        email, hashedPass, salt, namaUser, new Date()
      ]);
    } else {
      namaUser = existingUser.name;
    }
    
    // Tambahkan ke sheet Members sebagai Pending
    memSheet.appendRow([
      "m" + Date.now(), data.trip_id || "", email, namaUser, data.role || "member", "pending"
    ]);

    // Kirim Notifikasi Internal (Alerts)
    createAlert({
      email: email,
      title: "Undangan Trip Baru",
      message: `Anda diundang ke trip "${data.trip_name || 'Trip Baru'}" oleh ${data.inviterName || 'teman'}.`,
      type: "invite",
      trip_id: data.trip_id
    });

    // KIRIM EMAIL
    let emailSent = true;
    let emailError = "";
    try {
      const appUrl = "https://script.google.com/macros/s/AKfycbwsdro0eN-dZcZY70wIYUDQphV7u_pISaeM0lzzcVfpQ4z-brBFhd37j5NSH2eIhLOmpw/exec"; 
      let subject = `Undangan Kolaborasi Trip: ${data.trip_name}`;
      let body = `Halo ${namaUser},\n\n`;
      body += `${data.inviterName} mengundang Anda untuk bergabung dalam rencana perjalanan "${data.trip_name}".\n\n`;
      
      if (isNewUser) {
        body += `Karena Anda belum memiliki akun, kami telah membuatkan akun sementara untuk Anda:\n`;
        body += `Email: ${email}\n`;
        body += `Password Sementara: ${tempPassword}\n\n`;
      }
      
      body += `Silakan akses aplikasi melalui link berikut untuk menerima undangan dan mulai berkolaborasi:\n`;
      body += `${appUrl}\n\n`;
      body += `Salam,\nTrip Planner App`;
      
      GmailApp.sendEmail(email, subject, body);
    } catch (mailErr) {
      emailSent = false;
      emailError = mailErr.message;
      console.error("Gagal kirim email: " + emailError);
    }

    let successMsg = isNewUser ? "User baru dibuat & undangan terkirim!" : "Undangan telah dikirim!";
    if (!emailSent) {
      successMsg += " (Catatan: Email gagal terkirim: " + emailError + ". Pastikan Anda telah memberikan izin kirim email di Google Script)";
    }

    return { 
      success: true, 
      message: successMsg 
    };
  } catch(e) { return { success: false, message: e.message }; }
}

function deleteMember(id) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Members");
    const target = getSheetDataAsObjects("Members").find(m => String(m.member_id) === String(id));
    if (target) {
      sheet.deleteRow(target._rowIndex);
      return { success: true, message: "Anggota dikeluarkan" };
    }
    return { success: false, message: "Gagal mengeluarkan" };
  } catch(e) { return { success: false, message: e.message }; }
}
