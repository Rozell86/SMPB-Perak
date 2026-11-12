// =========================================================================
// BACKEND UTAMA: SISTEM MATRIKS PEMBELAJARAN BERSEPADU (SMPB)  
// =========================================================================

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Sistem Matrik Pembelajaran Bersepadu Negeri Perak')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}

// =========================================================================
// FUNGSI HELPER: VERSI TURBO (KALIS RALAT EJAAN & NAMA SHEET)
// =========================================================================
function ambilSheetKalisRalat(ss, namaSasaran) {
  var sheetPantas = ss.getSheetByName(namaSasaran);
  if (sheetPantas) return sheetPantas;

  var sheets = ss.getSheets();
  var sasaranLower = namaSasaran.trim().toLowerCase();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().trim().toLowerCase() === sasaranLower) {
      return sheets[i];
    }
  }
  throw new Error("Helaian (Sheet) bernama '" + namaSasaran + "' tidak dijumpai!");
}

// =========================================================================
// 1. PROSES LOG MASUK BAHARU & SEMAKAN KELAYAKAN (PETA SHEET 'user')
// =========================================================================
function checkUserCredentials(userId, password) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ambilSheetKalisRalat(ss, 'user');
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return h.toString().trim().toUpperCase(); });

    var colUser = headers.indexOf('USERNAME');
    var colPass = headers.indexOf('PASSWORD') !== -1 ? headers.indexOf('PASSWORD') : headers.indexOf('PASSSWORD');
    var colRole = headers.indexOf('ROLE');
    var colStatus = headers.indexOf('STATUS_TUKAR');
    var colEmail = headers.indexOf('EMEL');
    var colDaerah = headers.indexOf('DAERAH');
    var colKodDaerah = headers.indexOf('KOD_DAERAH');

    if (colUser === -1 || colPass === -1 || colRole === -1) {
      return { success: false, message: 'Tajuk lajur USERNAME, PASSWORD, atau ROLE tidak dijumpai.' };
    }

    var inputUser = userId.toString().trim().toLowerCase();
    var inputPass = password.toString().trim();

    for (var i = 1; i < data.length; i++) {
      var dbUser = data[i][colUser] ? data[i][colUser].toString().trim() : "";
      var dbEmail = (colEmail !== -1 && data[i][colEmail]) ? data[i][colEmail].toString().trim() : "";
      var dbPass = data[i][colPass] ? data[i][colPass].toString().trim() : "";

      var matchUser = (dbUser.toLowerCase() === inputUser) || (dbEmail.toLowerCase() === inputUser);

      if (matchUser && dbPass === inputPass) {
        var dbRole = data[i][colRole] ? data[i][colRole].toString().trim().toUpperCase() : "";
        var dbStatus = (colStatus !== -1 && data[i][colStatus]) ? data[i][colStatus].toString().trim().toUpperCase() : "SUDAH";
        var dbDaerah = (colDaerah !== -1 && data[i][colDaerah]) ? data[i][colDaerah].toString().trim() : "";
        var dbKodDaerah = (colKodDaerah !== -1 && data[i][colKodDaerah]) ? data[i][colKodDaerah].toString().trim() : "";

        var kodSekolahUser = (dbRole === 'SEKOLAH') ? dbUser.toUpperCase() : "";
        var dataSistem = dapatkanProfilSekolahDirect(kodSekolahUser);

        return {
          success: true,
          userId: dbUser.toUpperCase(),
          username: dbUser.toUpperCase(),
          emel: dbEmail,
          daerah: dbDaerah,
          kod_daerah: dbKodDaerah,
          role: dbRole,
          statusTukar: dbStatus,
          kod_sekolah: kodSekolahUser,
          dataSistem: dataSistem
        };
      }
    }

    return { success: false, message: 'ID Pengguna atau Kata Laluan tidak sah.' };

  } catch (error) {
    return { success: false, message: 'Ralat Server: ' + error.message };
  }
}

// =========================================================================
// 1B. KEMASKINI KATA LALUAN & STATUS_TUKAR (CARA A - KHUSUS SEKOLAH)
// =========================================================================
function updatePasswordCaraA(userId, newPassword) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ambilSheetKalisRalat(ss, 'user');
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return h.toString().trim().toUpperCase(); });

    var colUser = headers.indexOf('USERNAME');
    var colPass = headers.indexOf('PASSWORD');
    var colStatus = headers.indexOf('STATUS_TUKAR');
    var colRole = headers.indexOf('ROLE');
    var colEmail = headers.indexOf('EMEL');
    var colDaerah = headers.indexOf('DAERAH');
    var colKodDaerah = headers.indexOf('KOD_DAERAH');

    for (var i = 1; i < data.length; i++) {
      var dbUser = data[i][colUser] ? data[i][colUser].toString().trim() : "";
      if (dbUser.toUpperCase() === userId.toString().trim().toUpperCase()) {
        var rowNum = i + 1;
        
        sheet.getRange(rowNum, colPass + 1).setValue(newPassword);
        if (colStatus !== -1) {
          sheet.getRange(rowNum, colStatus + 1).setValue('SUDAH');
        }

        SpreadsheetApp.flush();

        var dbRole = data[i][colRole] ? data[i][colRole].toString().trim().toUpperCase() : "";
        var dbEmail = (colEmail !== -1 && data[i][colEmail]) ? data[i][colEmail].toString().trim() : "";
        var dbDaerah = (colDaerah !== -1 && data[i][colDaerah]) ? data[i][colDaerah].toString().trim() : "";
        var dbKodDaerah = (colKodDaerah !== -1 && data[i][colKodDaerah]) ? data[i][colKodDaerah].toString().trim() : "";

        var kodSekolahUser = (dbRole === 'SEKOLAH') ? dbUser.toUpperCase() : "";
        var dataSistem = dapatkanProfilSekolahDirect(kodSekolahUser);

        return {
          success: true,
          userData: {
            userId: dbUser.toUpperCase(),
            username: dbUser.toUpperCase(),
            emel: dbEmail,
            daerah: dbDaerah,
            kod_daerah: dbKodDaerah,
            role: dbRole,
            statusTukar: 'SUDAH',
            kod_sekolah: kodSekolahUser,
            dataSistem: dataSistem
          }
        };
      }
    }

    return { success: false, message: 'Rekod pengguna tidak dijumpai.' };

  } catch (error) {
    return { success: false, message: 'Gagal mengemaskini kata laluan: ' + error.message };
  }
}

// =========================================================================
// 1C. MUATKAN MEMORI / LAMAN APLIKASI UTAMA
// =========================================================================
function loadMainPage() {
  return HtmlService.createHtmlOutputFromFile('App').getContent(); 
}

// =========================================================================
// 1D. PROSES LOG MASUK WARISAN (BACKWARD COMPATIBILITY)
// =========================================================================
function prosesLogin(emailInput, password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetUser = ambilSheetKalisRalat(ss, 'user');
  var dataUser = sheetUser.getDataRange().getValues();
  var headers = dataUser[0].map(function(h) { return h.toString().trim().toUpperCase(); });

  var colUser = headers.indexOf('USERNAME');
  var colPass = headers.indexOf('PASSWORD') !== -1 ? headers.indexOf('PASSWORD') : headers.indexOf('PASSSWORD');
  var colEmail = headers.indexOf('EMEL');
  var colRole = headers.indexOf('ROLE');
  var colStatus = headers.indexOf('STATUS_TUKAR');
  var colDaerah = headers.indexOf('DAERAH');
  var colKodDaerah = headers.indexOf('KOD_DAERAH');

  var inputEmail = emailInput.toString().trim().toLowerCase();
  var inputPass = password.toString().trim();

  var loginBerjaya = false;
  var userSession = {};

  for (var i = 1; i < dataUser.length; i++) {
    var dbUser = (colUser !== -1 && dataUser[i][colUser]) ? dataUser[i][colUser].toString().trim() : (dataUser[i][1] ? dataUser[i][1].toString().trim() : "");
    var dbPass = (colPass !== -1 && dataUser[i][colPass]) ? dataUser[i][colPass].toString().trim() : (dataUser[i][2] ? dataUser[i][2].toString().trim() : "");
    var dbEmail = (colEmail !== -1 && dataUser[i][colEmail]) ? dataUser[i][colEmail].toString().trim().toLowerCase() : (dataUser[i][3] ? dataUser[i][3].toString().trim().toLowerCase() : "");

    var matchUser = (dbEmail === inputEmail) || (dbUser.toLowerCase() === inputEmail);

    if (matchUser && dbPass === inputPass) {
      loginBerjaya = true;
      var dbRole = (colRole !== -1 && dataUser[i][colRole]) ? dataUser[i][colRole].toString().trim().toUpperCase() : (dataUser[i][6] ? dataUser[i][6].toString().trim().toUpperCase() : "");
      var dbStatus = (colStatus !== -1 && dataUser[i][colStatus]) ? dataUser[i][colStatus].toString().trim().toUpperCase() : "SUDAH";
      var dbDaerah = (colDaerah !== -1 && dataUser[i][colDaerah]) ? dataUser[i][colDaerah].toString().trim() : (dataUser[i][4] ? dataUser[i][4].toString().trim() : "");
      var dbKodDaerah = (colKodDaerah !== -1 && dataUser[i][colKodDaerah]) ? dataUser[i][colKodDaerah].toString().trim() : (dataUser[i][5] ? dataUser[i][5].toString().trim() : "");

      var kodSekolahUser = "";
      if (dbRole === 'SEKOLAH') {
        kodSekolahUser = dbUser.toUpperCase();
      }

      userSession = {
        username: dbUser.toUpperCase(),
        emel: dbEmail,
        daerah: dbDaerah,
        kod_daerah: dbKodDaerah,
        role: dbRole,
        statusTukar: dbStatus,
        kod_sekolah: kodSekolahUser
      };
      break;
    }
  }

  if (!loginBerjaya) {
    return { status: "FAILED", message: "Emel Pengguna / ID atau Kata Laluan Salah!" };
  }

  var dataSistem = dapatkanProfilSekolahDirect(userSession.kod_sekolah);
  return { status: "SUCCESS", user: userSession, dataSistem: dataSistem };
}

function kemaskiniKataLaluan(usernameOrEmail, newPassword) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetUser = ambilSheetKalisRalat(ss, 'user');
    var dataUser = sheetUser.getDataRange().getValues();
    var headers = dataUser[0].map(function(h) { return h.toString().trim().toUpperCase(); });

    var colUser = headers.indexOf('USERNAME');
    var colPass = headers.indexOf('PASSWORD') !== -1 ? headers.indexOf('PASSWORD') : headers.indexOf('PASSSWORD');
    var colEmail = headers.indexOf('EMEL');
    var colStatus = headers.indexOf('STATUS_TUKAR');

    var inputTarget = usernameOrEmail.toString().trim().toLowerCase();

    for (var i = 1; i < dataUser.length; i++) {
      var dbUser = (colUser !== -1 && dataUser[i][colUser]) ? dataUser[i][colUser].toString().trim().toLowerCase() : "";
      var dbEmail = (colEmail !== -1 && dataUser[i][colEmail]) ? dataUser[i][colEmail].toString().trim().toLowerCase() : "";

      if (dbUser === inputTarget || dbEmail === inputTarget) {
        var rowNum = i + 1;
        
        if (colPass !== -1) {
          sheetUser.getRange(rowNum, colPass + 1).setValue(newPassword.trim());
        } else {
          sheetUser.getRange(rowNum, 3).setValue(newPassword.trim());
        }
        
        if (colStatus !== -1) {
          sheetUser.getRange(rowNum, colStatus + 1).setValue("SUDAH");
        }

        SpreadsheetApp.flush();
        return prosesLogin(usernameOrEmail, newPassword.trim());
      }
    }
    return { status: "FAILED", message: "Akaun pengguna tidak dijumpai." };
  } catch (e) {
    return { status: "FAILED", message: "Ralat Server: " + e.message };
  }
}

// =========================================================================
// 2. FUNGSI KHUSUS UNTUK MEMANGGIL DATA PROFIL SEGAR (KEMASKINI STATUS LOKASI)
// =========================================================================
function dapatkanProfilSekolahDirect(kodSekolah) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var kodSek = kodSekolah.trim().toUpperCase();
  
  var dataSistem = {
    profilSekolah: null,
    kpmSekolahList: [],
    pmpSekolahList: [],
    senaraiMp: [],
    tetapan: {} 
  };

  try {
    var sheetTetapan = ss.getSheetByName('tb_tetapan');
    if (sheetTetapan) {
      var dataTetapan = sheetTetapan.getDataRange().getValues();
      for (var t = 1; t < dataTetapan.length; t++) {
        dataSistem.tetapan[dataTetapan[t][0]] = dataTetapan[t][1];
      }
    }
  } catch(e) {}

  try {
    var sheetMp = ambilSheetKalisRalat(ss, 'tbmp');
    var dataMp = sheetMp.getDataRange().getValues();
    for (var m = 1; m < dataMp.length; m++) {
      dataSistem.senaraiMp.push({
        kod: dataMp[m][1] ? dataMp[m][1].toString().trim() : "",
        nama: dataMp[m][2] ? dataMp[m][2].toString().trim() : "",
        kategori: dataMp[m][3] ? dataMp[m][3].toString().trim().toUpperCase() : "UMUM" 
      });
    }
  } catch(e) { Logger.log("Ralat tbmp: " + e.message); }

  try {
    var sheetSek = ambilSheetKalisRalat(ss, 'tbsekolah');
    var dataSek = sheetSek.getDataRange().getValues();
    
    // SIFU TAMBAH: Pastikan kita tangkap 'STATUS LOKASI' di Lajur 27 (Z) atau 28 (AA)
    for (var j = 1; j < dataSek.length; j++) {
      var dbKodSek = dataSek[j][5] ? dataSek[j][5].toString().trim().toUpperCase() : ""; 
      if (dbKodSek === kodSek) {
        dataSistem.profilSekolah = {
          rowNum: j + 1, 
          kategori_sekolah: dataSek[j][1] ? dataSek[j][1].toString().trim().toUpperCase() : "RENDAH",
          ppd: dataSek[j][3] ? dataSek[j][3].toString().trim() : "",      
          daerah: dataSek[j][4] ? dataSek[j][4].toString().trim() : "",   
          kod: dataSek[j][5].toString().trim(),                                          
          nama: dataSek[j][6].toString().trim(),                               
          alamat: dataSek[j][7] ? dataSek[j][7].toString().trim() : "",   
          poskod: dataSek[j][8] ? dataSek[j][8].toString().trim() : "",   
          bandar: dataSek[j][9] ? dataSek[j][9].toString().trim() : "",
          tel_sekolah: dataSek[j][10] ? dataSek[j][10].toString().trim() : "",
          emel_sekolah: dataSek[j][11] ? dataSek[j][11].toString().trim() : "",
          moeis: dataSek[j][12] ? dataSek[j][12].toString().trim() : "",
          nama_gb: dataSek[j][13] ? dataSek[j][13].toString().trim() : "",   
          kp_gb: dataSek[j][14] ? dataSek[j][14].toString().trim() : "",     
          tel_gb: dataSek[j][15] ? dataSek[j][15].toString().trim() : "",    
          emel_gb: dataSek[j][16] ? dataSek[j][16].toString().trim() : "",   
          nama_sup: dataSek[j][17] ? dataSek[j][17].toString().trim() : "",  
          kp_sup: dataSek[j][18] ? dataSek[j][18].toString().trim() : "",   
          jantina_sup: dataSek[j][19] ? dataSek[j][19].toString().trim() : "", 
          tel_sup: dataSek[j][20] ? dataSek[j][20].toString().trim() : "",  
          emel_sup: dataSek[j][21] ? dataSek[j][21].toString().trim() : "", 
          bil_kelas: dataSek[j][22] ? dataSek[j][22].toString().trim() : "0", 
          jum_calon: dataSek[j][23] ? dataSek[j][23].toString().trim() : "0", 
          pengesahan_selesai: dataSek[j][24] ? dataSek[j][24].toString().trim() : "",
          // TANGKAP STATUS LOKASI (LAJUR AA/Z)
          status_lokasi: dataSek[j][26] ? dataSek[j][26].toString().trim().toUpperCase() : "" 
        };
        break;
      }
    }
  } catch (e) { Logger.log("Ralat tbsekolah: " + e.message); }

  try {
    var sheetKpmp = ambilSheetKalisRalat(ss, 'tbkpmp');
    var dataKpmp = sheetKpmp.getDataRange().getValues();
    for (var k = 1; k < dataKpmp.length; k++) {
      var kSek = dataKpmp[k][1] ? dataKpmp[k][1].toString().trim().toUpperCase() : "";
      if (kSek === kodSek) {
        dataSistem.kpmSekolahList.push({
          rowNum: k + 1,
          ppd: dataKpmp[k][2] ? dataKpmp[k][2].toString().trim() : "",    
          kod_mp: dataKpmp[k][3] ? dataKpmp[k][3].toString().trim() : "", 
          nama_kpmp: dataKpmp[k][4] ? dataKpmp[k][4].toString().trim() : "", 
          no_kp: dataKpmp[k][5] ? dataKpmp[k][5].toString().trim() : "",    
          jantina: dataKpmp[k][6] ? dataKpmp[k][6].toString().trim() : "",  
          no_tel: dataKpmp[k][7] ? dataKpmp[k][7].toString().trim() : "",   
          emel: dataKpmp[k][8] ? dataKpmp[k][8].toString().trim() : ""      
        });
      }
    }
  } catch(e) { Logger.log("Ralat tbkpmp: " + e.message); }

  try {
    var sheetPmp = ambilSheetKalisRalat(ss, 'tbpmp');
    var dataPmp = sheetPmp.getDataRange().getValues();
    for (var p = 1; p < dataPmp.length; p++) {
      var pSek = dataPmp[p][1] ? dataPmp[p][1].toString().trim().toUpperCase() : "";
      if (pSek === kodSek) {
        dataSistem.pmpSekolahList.push({
          rowNum: p + 1,
          ppd: dataPmp[p][2] ? dataPmp[p][2].toString().trim() : "",    
          kod_mp: dataPmp[p][3] ? dataPmp[p][3].toString().trim() : "", 
          nama_pmp: dataPmp[p][4] ? dataPmp[p][4].toString().trim() : "",  
          no_kp: dataPmp[p][5] ? dataPmp[p][5].toString().trim() : "",    
          jantina: dataPmp[p][6] ? dataPmp[p][6].toString().trim() : "",  
          no_tel: dataPmp[p][7] ? dataPmp[p][7].toString().trim() : "",   
          emel: dataPmp[p][8] ? dataPmp[p][8].toString().trim() : ""      
        });
      }
    }
  } catch(e) { Logger.log("Ralat tbpmp: " + e.message); }

  return dataSistem;
}

// 3. FUNGSI SIMPAN PROFIL SEKOLAH (TERMASUK BAWA JUM CALON)
function simpanProfilSekolah(rowNum, d) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ambilSheetKalisRalat(ss, 'tbsekolah');
  var r = Number(rowNum);
  
  var inputKpGb = d.kp_gb ? d.kp_gb.toString().trim() : "";
  var inputKpSup = d.kp_sup ? d.kp_sup.toString().trim() : "";
  
  var dataSek = sheet.getDataRange().getValues();
  for (var i = 1; i < dataSek.length; i++) {
    var rowDB = i + 1;
    if (rowDB === r) continue;
    var dbKpGb = dataSek[i][14] ? dataSek[i][14].toString().trim() : ""; 
    var dbKpSup = dataSek[i][18] ? dataSek[i][18].toString().trim() : ""; 
    
    var berlakuTarikan = false;
    
    if (inputKpGb !== "" && dbKpGb === inputKpGb) {
      sheet.getRange(rowDB, 14).setValue("");
      sheet.getRange(rowDB, 15).setValue("");
      sheet.getRange(rowDB, 16).setValue("");
      sheet.getRange(rowDB, 17).setValue("");
      berlakuTarikan = true;
    }
    
    if (inputKpSup !== "" && dbKpSup === inputKpSup) {
      sheet.getRange(rowDB, 18).setValue("");
      sheet.getRange(rowDB, 19).setValue("");
      sheet.getRange(rowDB, 20).setValue("");
      sheet.getRange(rowDB, 21).setValue("");
      sheet.getRange(rowDB, 22).setValue("");
      berlakuTarikan = true;
    }
    
    if (berlakuTarikan) {
      sheet.getRange(rowDB, 25).setValue(""); // Reset pengesahan profil
    }
  }

  sheet.getRange(r, 14).setValue(d.nama_gb);
  sheet.getRange(r, 15).setValue(d.kp_gb);         
  sheet.getRange(r, 16).setValue(d.tel_gb);
  sheet.getRange(r, 17).setValue(d.emel_gb);       
  sheet.getRange(r, 18).setValue(d.nama_sup);
  sheet.getRange(r, 19).setValue(d.kp_sup);        
  sheet.getRange(r, 20).setValue(d.jantina_sup);
  sheet.getRange(r, 21).setValue(d.tel_sup);       
  sheet.getRange(r, 22).setValue(d.emel_sup);
  sheet.getRange(r, 23).setValue(Number(d.bil_kelas)); 
  sheet.getRange(r, 24).setValue(Number(d.jum_calon));
  
  SpreadsheetApp.flush();
  
  var kodSekolah = sheet.getRange(r, 6).getValue().toString().trim();
  var dataSistemSegar = dapatkanProfilSekolahDirect(kodSekolah);
  return { status: "SUCCESS", message: "Profil Sekolah Berjaya Dikemaskini!", dataSistem: dataSistemSegar };
}

// 4. FUNGSI DATA KPMP
function urusKpmp(username, d, rowNum) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ambilSheetKalisRalat(ss, 'tbkpmp');
    var inputKp = d.no_kp.trim();
    var inputMp = d.kod_mp.trim();
    var currentSchool = username.trim().toUpperCase();
    if (rowNum) {
      var r = Number(rowNum);
      sheet.getRange(r, 5).setValue(d.nama_kpmp); 
      sheet.getRange(r, 6).setValue(inputKp);
      sheet.getRange(r, 7).setValue(d.jantina);   
      sheet.getRange(r, 8).setValue(d.no_tel);
      sheet.getRange(r, 9).setValue(d.emel);
    } else {
      var dataKpmp = sheet.getDataRange().getValues();
      var rowsToDelete = []; 
      
      for (var i = 1; i < dataKpmp.length; i++) {
        var dbSchool = dataKpmp[i][1] ? dataKpmp[i][1].toString().trim().toUpperCase() : "";
        var dbMp = dataKpmp[i][3] ? dataKpmp[i][3].toString().trim() : "";
        var dbKp = dataKpmp[i][5] ? dataKpmp[i][5].toString().trim() : "";
        if (dbSchool === currentSchool) {
          if (dbMp === inputMp) return { status: "FAILED", message: "Mata Pelajaran ini sudah mempunyai Ketua Pentaksir (KPMP)!" };
        } else {
          if (dbKp === inputKp && inputKp !== "") rowsToDelete.push(i + 1);
        }
      }
      if (rowsToDelete.length > 0) { rowsToDelete.reverse().forEach(function(rowDel) { sheet.deleteRow(rowDel); }); }
      var lastRow = sheet.getLastRow();
      var dataId = lastRow === 1 ? 1 : Number(sheet.getRange(lastRow, 1).getValue()) + 1;
      var dataSistemSekolah = dapatkanProfilSekolahDirect(currentSchool);
      var kodPpdSekolah = dataSistemSekolah.profilSekolah ? dataSistemSekolah.profilSekolah.ppd : "";
      sheet.appendRow([dataId, currentSchool, kodPpdSekolah, inputMp, d.nama_kpmp, inputKp, d.jantina, d.no_tel, d.emel]);
    }
    return { status: "SUCCESS" };
  } catch(e) { return { status: "FAILED", message: e.message }; }
}

// 5. FUNGSI DATA PMP
function urusPmp(username, d, rowNum) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ambilSheetKalisRalat(ss, 'tbpmp');
    var inputKp = d.no_kp.trim();
    var inputMp = d.kod_mp.trim(); 
    var currentSchool = username.trim().toUpperCase();
    if (rowNum) {
      var r = Number(rowNum);
      sheet.getRange(r, 5).setValue(d.nama_pmp);  
      sheet.getRange(r, 6).setValue(inputKp);
      sheet.getRange(r, 7).setValue(d.jantina);   
      sheet.getRange(r, 8).setValue(d.no_tel);
      sheet.getRange(r, 9).setValue(d.emel);
    } else {
      var dataPmp = sheet.getDataRange().getValues();
      var rowsToDelete = [];
      for (var i = 1; i < dataPmp.length; i++) {
        var dbSchool = dataPmp[i][1] ? dataPmp[i][1].toString().trim().toUpperCase() : "";
        var dbMp = dataPmp[i][3] ? dataPmp[i][3].toString().trim() : "";
        var dbKp = dataPmp[i][5] ? dataPmp[i][5].toString().trim() : "";
        if (dbSchool === currentSchool) {
          if (dbMp === inputMp && dbKp === inputKp) return { status: "FAILED", message: "Guru ini telah pun didaftarkan sebagai Pentaksir untuk mata pelajaran ini!" };
        } else {
          if (dbKp === inputKp && inputKp !== "") rowsToDelete.push(i + 1);
        }
      }
      if (rowsToDelete.length > 0) { rowsToDelete.reverse().forEach(function(rowDel) { sheet.deleteRow(rowDel); }); }
      var lastRow = sheet.getLastRow();
      var dataId = lastRow === 1 ? 1 : Number(sheet.getRange(lastRow, 1).getValue()) + 1;
      var dataSistemSekolah = dapatkanProfilSekolahDirect(currentSchool);
      var kodPpdSekolah = dataSistemSekolah.profilSekolah ? dataSistemSekolah.profilSekolah.ppd : "";
      sheet.appendRow([dataId, currentSchool, kodPpdSekolah, inputMp, d.nama_pmp, inputKp, d.jantina, d.no_tel, d.emel]);
    }
    return { status: "SUCCESS" };
  } catch(e) { return { status: "FAILED", message: e.message }; }
}

// 6. FUNGSI HAPUS REKOD
function padamRekod(sheetName, rowNum) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ambilSheetKalisRalat(ss, sheetName);
    sheet.deleteRow(Number(rowNum));
    return { status: "SUCCESS" };
  } catch(e) { return { status: "FAILED", message: e.message }; }
}

// =========================================================================
// 7. FUNGSI KHUSUS UNTUK MEMANGGIL DATA KESELURUHAN (ROLE: PPD / JPN)
// =========================================================================
function mendapatkanDataPPD_JPN(role, kodDaerah, username) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dataSistem = {
      senaraiSekolah: [], senaraiKpmp: [], senaraiPmp: [], senaraiMp: [],
      senaraiDaerah: [], senaraiMatrik: [], senaraiBilikKebal: [], 
      senaraiKp: [], senaraiPk: [], senaraiKpnKpkw: [], senaraiPgws: [], 
      tetapan: {} 
    };
    
    var roleSemasa = role ? role.toString().trim().toUpperCase() : "";
    var filterKodDaerah = "";
    
    if (username && username.toString().trim() !== "") filterKodDaerah = username.toString().trim().toUpperCase();
    else if (kodDaerah && kodDaerah.toString().trim() !== "") filterKodDaerah = kodDaerah.toString().trim().toUpperCase();

    if (roleSemasa === 'PPD' && filterKodDaerah === "") {
      try {
        var userEmailAktif = Session.getActiveUser().getEmail().toLowerCase().trim();
        var sheetUser = ambilSheetKalisRalat(ss, 'user');
        var dataUser = sheetUser.getDataRange().getValues();
        for (var u = 1; u < dataUser.length; u++) {
          if ((dataUser[u][3] || "").toString().trim().toLowerCase() === userEmailAktif) {
            filterKodDaerah = dataUser[u][5] ? dataUser[u][5].toString().trim().toUpperCase() : ""; break;
          }
        }
      } catch(errUser) {}
    }
    if (roleSemasa === 'PPD' && filterKodDaerah === "") filterKodDaerah = "A080";
    filterKodDaerah = filterKodDaerah.trim().toUpperCase();

    // TETAPAN (MASTER SWITCH)
    try {
      var sheetTetapan = ss.getSheetByName('tb_tetapan');
      if (sheetTetapan) {
        var dataTetapan = sheetTetapan.getDataRange().getValues();
        for (var t = 1; t < dataTetapan.length; t++) {
          dataSistem.tetapan[dataTetapan[t][0]] = dataTetapan[t][1];
        }
      }
    } catch(e) {}

    try { var dDae = ambilSheetKalisRalat(ss, 'tbdaerah').getDataRange().getValues(); for (var d = 1; d < dDae.length; d++) dataSistem.senaraiDaerah.push({ ppd: dDae[d][5] ? dDae[d][5].toString().trim().toUpperCase() : "", nama: dDae[d][2] ? dDae[d][2].toString().trim().toUpperCase() : "" }); } catch(e) {}
    try { var dMp = ambilSheetKalisRalat(ss, 'tbmp').getDataRange().getValues(); for (var m = 1; m < dMp.length; m++) dataSistem.senaraiMp.push({ kod: dMp[m][1] ? dMp[m][1].toString().trim() : "", nama: dMp[m][2] ? dMp[m][2].toString().trim() : "", kategori: dMp[m][3] ? dMp[m][3].toString().trim().toUpperCase() : "UMUM" }); } catch(e) {}
    
    // SEKOLAH
    try { 
      var dSek = ambilSheetKalisRalat(ss, 'tbsekolah').getDataRange().getValues(); 
      for (var j = 1; j < dSek.length; j++) { 
        var dbPpd = dSek[j][3] ? dSek[j][3].toString().trim().toUpperCase() : ""; 
        if (roleSemasa === 'PPD' && dbPpd !== filterKodDaerah) continue; 
        
        dataSistem.senaraiSekolah.push({ 
          rowNum: j + 1, kategori_sekolah: dSek[j][1] ? dSek[j][1].toString().trim().toUpperCase() : "RENDAH", 
          ppd: dbPpd, daerah: dSek[j][4] ? dSek[j][4].toString().trim() : "", kod: dSek[j][5] ? dSek[j][5].toString().trim() : "", 
          nama: dSek[j][6] ? dSek[j][6].toString().trim() : "", alamat: dSek[j][7] ? dSek[j][7].toString().trim() : "", 
          poskod: dSek[j][8] ? dSek[j][8].toString().trim() : "", bandar: dSek[j][9] ? dSek[j][9].toString().trim() : "",
          tel_sekolah: dSek[j][10] ? dSek[j][10].toString().trim() : "", emel_sekolah: dSek[j][11] ? dSek[j][11].toString().trim() : "",
          nama_gb: dSek[j][13] ? dSek[j][13].toString().trim() : "", kp_gb: dSek[j][14] ? dSek[j][14].toString().trim() : "",
          tel_gb: dSek[j][15] ? dSek[j][15].toString().trim() : "", emel_gb: dSek[j][16] ? dSek[j][16].toString().trim() : "",
          nama_sup: dSek[j][17] ? dSek[j][17].toString().trim() : "", kp_sup: dSek[j][18] ? dSek[j][18].toString().trim() : "",
          jantina_sup: dSek[j][19] ? dSek[j][19].toString().trim() : "", tel_sup: dSek[j][20] ? dSek[j][20].toString().trim() : "",
          emel_sup: dSek[j][21] ? dSek[j][21].toString().trim() : "", bil_kelas: dSek[j][22] ? dSek[j][22].toString().trim() : "0",
          jum_calon: dSek[j][23] ? dSek[j][23].toString().trim() : "0", pengesahan_selesai: dSek[j][24] ? dSek[j][24].toString().trim().toUpperCase() : "",
          pengesahan_pengawas: dSek[j][25] ? dSek[j][25].toString().trim().toUpperCase() : "",
          status_lokasi: dSek[j][26] ? dSek[j][26].toString().trim().toUpperCase() : "" // TANGKAP LOKASI (KEMASKINI)
        }); 
      } 
    } catch(e) {}

    // KPMP
    try { 
      var dKpmp = ambilSheetKalisRalat(ss, 'tbkpmp').getDataRange().getValues(); 
      for (var k = 1; k < dKpmp.length; k++) { 
        var dbPpdKpmp = dKpmp[k][2] ? dKpmp[k][2].toString().trim().toUpperCase() : "";
        if (roleSemasa === 'PPD' && dbPpdKpmp !== filterKodDaerah) continue;
        dataSistem.senaraiKpmp.push({ kod_sekolah: dKpmp[k][1] ? dKpmp[k][1].toString().trim() : "", ppd: dbPpdKpmp, kod_mp: dKpmp[k][3] ? dKpmp[k][3].toString().trim() : "", nama_kpmp: dKpmp[k][4] ? dKpmp[k][4].toString().trim() : "", no_kp: dKpmp[k][5] ? dKpmp[k][5].toString().trim() : "", no_tel: dKpmp[k][7] ? dKpmp[k][7].toString().trim() : "", emel: dKpmp[k][8] ? dKpmp[k][8].toString().trim() : "" }); 
      } 
    } catch(e) {}
    
    // PMP
    try { 
      var dPmp = ambilSheetKalisRalat(ss, 'tbpmp').getDataRange().getValues(); 
      for (var p = 1; p < dPmp.length; p++) { 
        var dbPpdPmp = dPmp[p][2] ? dPmp[p][2].toString().trim().toUpperCase() : "";
        if (roleSemasa === 'PPD' && dbPpdPmp !== filterKodDaerah) continue;
        dataSistem.senaraiPmp.push({ kod_sekolah: dPmp[p][1] ? dPmp[p][1].toString().trim() : "", ppd: dbPpdPmp, kod_mp: dPmp[p][3] ? dPmp[p][3].toString().trim() : "", nama_pmp: dPmp[p][4] ? dPmp[p][4].toString().trim() : "", no_kp: dPmp[p][5] ? dPmp[p][5].toString().trim() : "", no_tel: dPmp[p][7] ? dPmp[p][7].toString().trim() : "", emel: dPmp[p][8] ? dPmp[p][8].toString().trim() : "" }); 
      } 
    } catch(e) {}
    
    // MATRIK OPERASI
    try { 
      var dMat = ss.getSheetByName('tbmatrik_operasi').getDataRange().getValues(); 
      for (var mo = 1; mo < dMat.length; mo++) {
        var dbDaeMat = dMat[mo][3] ? dMat[mo][3].toString().trim().toUpperCase() : "";
        if (roleSemasa === 'PPD' && dbDaeMat !== filterKodDaerah) continue;
        dataSistem.senaraiMatrik.push({ tahun: dMat[mo][1] ? dMat[mo][1].toString().trim() : "", kod_sekolah: dMat[mo][2] ? dMat[mo][2].toString().trim().toUpperCase() : "", matrik_tahap: dMat[mo][4] ? dMat[mo][4].toString().trim().toUpperCase() : "", id_bilik_kebal: dMat[mo][5] ? dMat[mo][5].toString().trim() : "", no_kp_pk1: dMat[mo][6] ? dMat[mo][6].toString().trim() : "", no_kp_pk2: dMat[mo][7] ? dMat[mo][7].toString().trim() : "" }); 
      }
    } catch(e) {}
    
    // BILIK KEBAL
    try { 
      var dBk = ss.getSheetByName('tbbilikkebal').getDataRange().getValues(); 
      for (var bk = 1; bk < dBk.length; bk++) {
        var dbDaeBk = dBk[bk][2] ? dBk[bk][2].toString().trim().toUpperCase() : "";
        if (roleSemasa === 'PPD' && dbDaeBk !== filterKodDaerah) continue;
        dataSistem.senaraiBilikKebal.push({ id_bilik_kebal: dBk[bk][0] ? dBk[bk][0].toString().trim() : "", nama_bilik_kebal: dBk[bk][1] ? dBk[bk][1].toString().trim() : "", daerah: dbDaeBk }); 
      }
    } catch(e) {}
    
    // KETUA PENGAWAS - SIFU FIX: Masukkan rowNum ke dalam array!
    try { 
      var dKp = ss.getSheetByName('tbkp').getDataRange().getValues(); 
      for (var kp = 1; kp < dKp.length; kp++) { 
        var dbDaerahKp = dKp[kp][3] ? dKp[kp][3].toString().trim().toUpperCase() : ""; 
        if (roleSemasa === 'PPD' && dbDaerahKp !== filterKodDaerah) continue; 
        
        dataSistem.senaraiKp.push({ 
          rowNum: kp + 1, // FIX PENTING!
          tahun: dKp[kp][1] ? dKp[kp][1].toString().trim() : "", 
          kod_sek_asal: dKp[kp][2] ? dKp[kp][2].toString().trim().toUpperCase() : "", 
          matrik_tahap: dKp[kp][4] ? dKp[kp][4].toString().trim().toUpperCase() : "", 
          nama_guru: dKp[kp][6] ? dKp[kp][6].toString().trim() : "", 
          no_kp: dKp[kp][7] ? dKp[kp][7].toString().trim() : "", 
          no_tel: dKp[kp][8] ? dKp[kp][8].toString().trim() : "", 
          emel: dKp[kp][9] ? dKp[kp][9].toString().trim() : "", 
          alamat_rumah: dKp[kp][10] ? dKp[kp][10].toString().trim().toUpperCase() : "", 
          kod_sek_cadangan: dKp[kp][11] ? dKp[kp][11].toString().trim().toUpperCase() : "", 
          kod_sek_tetap: dKp[kp][12] ? dKp[kp][12].toString().trim().toUpperCase() : "", 
          status_penetapan: dKp[kp][13] ? dKp[kp][13].toString().trim().toUpperCase() : "" 
        }); 
      } 
    } catch(e) {}
    
    // PENYELIA KAWASAN
    try { 
      var dPk = ss.getSheetByName('tbpk').getDataRange().getValues(); 
      for (var pk = 1; pk < dPk.length; pk++) { 
        var dbDaerahPk = dPk[pk][3] ? dPk[pk][3].toString().trim().toUpperCase() : ""; 
        if (roleSemasa === 'PPD' && dbDaerahPk !== filterKodDaerah) continue; 
        dataSistem.senaraiPk.push({ rowNum: pk + 1, id: dPk[pk][0], tahun: dPk[pk][1] ? dPk[pk][1].toString().trim() : "", kod_sek_asal: dPk[pk][2] ? dPk[pk][2].toString().trim().toUpperCase() : "", daerah: dbDaerahPk, matrik_tahap: dPk[pk][4] ? dPk[pk][4].toString().trim().toUpperCase() : "", jawatan_pk: dPk[pk][6] ? dPk[pk][6].toString().trim().toUpperCase() : "", nama_pk: dPk[pk][7] ? dPk[pk][7].toString().trim() : "", no_kp: dPk[pk][8] ? dPk[pk][8].toString().trim() : "", no_tel: dPk[pk][9] ? dPk[pk][9].toString().trim() : "", emel: dPk[pk][10] ? dPk[pk][10].toString().trim() : "", alamat_rumah: dPk[pk][11] ? dPk[pk][11].toString().trim().toUpperCase() : "" }); 
      } 
    } catch(e) {}

    // PENGAWAS
    try { 
      var dPgws = ss.getSheetByName('tbp').getDataRange().getValues(); 
      for (var p = 1; p < dPgws.length; p++) { 
        var dbDaerahPgws = dPgws[p][3] ? dPgws[p][3].toString().trim().toUpperCase() : ""; 
        if (roleSemasa === 'PPD' && dbDaerahPgws !== filterKodDaerah) continue; 
        dataSistem.senaraiPgws.push({ 
          rowNum: p + 1, id: dPgws[p][0], tahun: dPgws[p][1] ? dPgws[p][1].toString().trim() : "", 
          kod_sekolah: dPgws[p][2] ? dPgws[p][2].toString().trim().toUpperCase() : "", daerah: dbDaerahPgws, 
          matrik_tahap: dPgws[p][4] ? dPgws[p][4].toString().trim().toUpperCase() : "", nama: dPgws[p][5] ? dPgws[p][5].toString().trim() : "", 
          no_kp: dPgws[p][6] ? dPgws[p][6].toString().trim() : "", jantina: dPgws[p][7] ? dPgws[p][7].toString().trim() : "", 
          no_tel: dPgws[p][8] ? dPgws[p][8].toString().trim() : "", emel: dPgws[p][9] ? dPgws[p][9].toString().trim() : "" 
        }); 
      } 
    } catch(e) {}
    
    // KPN / KPKW 
    try {
      var sheetKpn = ss.getSheetByName('tbkpn_kpkw');
      if(sheetKpn) {
        var dataKpn = sheetKpn.getDataRange().getValues();
        for (var n = 1; n < dataKpn.length; n++) {
          var dbDaeKpn = dataKpn[n][4] ? dataKpn[n][4].toString().trim().toUpperCase() : "";
          if (roleSemasa === 'PPD' && dbDaeKpn !== filterKodDaerah) continue;
          dataSistem.senaraiKpnKpkw.push({
            rowNum: n + 1, 
            id: dataKpn[n][0], 
            tahun: dataKpn[n][1] ? dataKpn[n][1].toString().trim() : "", 
            matrik_tahap: dataKpn[n][2] ? dataKpn[n][2].toString().trim().toUpperCase() : "",
            peringkat: dataKpn[n][3] ? dataKpn[n][3].toString().trim().toUpperCase() : "", 
            daerah: dbDaeKpn, 
            kod_mp: dataKpn[n][5] ? dataKpn[n][5].toString().trim().toUpperCase() : "",
            nama_guru: dataKpn[n][6] ? dataKpn[n][6].toString().trim().toUpperCase() : "", 
            no_kp: dataKpn[n][7] ? dataKpn[n][7].toString().trim() : "",
            kod_sekolah: dataKpn[n][8] ? dataKpn[n][8].toString().trim().toUpperCase() : "", 
            no_tel: dataKpn[n][9] ? dataKpn[n][9].toString().trim() : "",
            emel: dataKpn[n][10] ? dataKpn[n][10].toString().trim() : "", 
            no_rujukan_surat: dataKpn[n][11] ? dataKpn[n][11].toString().trim() : ""
          });
        }
      }
    } catch(e) {}

    return { status: "SUCCESS", data: dataSistem };
  } catch(e) { return { status: "FAILED", message: e.message }; }
}

// =========================================================================
// KEMASKINI: FUNGSI SIMPAN MASTER SWITCH JADUAL KERJA (HANYA UNTUK JPN/ADMIN)
// =========================================================================
function simpanMasterSwitchDB(kunci, status) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("tb_tetapan");
    if (!sheet) {
      sheet = ss.insertSheet("tb_tetapan");
      sheet.appendRow(["kunci", "nilai"]);
    }
    var data = sheet.getDataRange().getValues();
    var wujud = false;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === kunci) {
        sheet.getRange(i + 1, 2).setValue(status);
        wujud = true;
        break;
      }
    }
    if (!wujud) {
      sheet.appendRow([kunci, status]);
    }
    SpreadsheetApp.flush();
    return { status: "SUCCESS" };
  } catch (e) {
    return { status: "ERROR", message: e.message };
  }
}

// =========================================================================
// 8. MODUL KETUA PENGAWAS (PERANAN SEKOLAH) - KEMASKINI PEDALAMAN
// =========================================================================
function dapatkanSekolahDaerahSama(kodPpdSemasa, kodSekAsal) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ambilSheetKalisRalat(ss, 'tbsekolah');
  var data = sheet.getDataRange().getValues();
  
  var tahunSemasa = new Date().getFullYear().toString();
  var sheetKp = ss.getSheetByName('tbkp');
  var kiraanCadangan = {};
  
  if (sheetKp) {
    var dataKp = sheetKp.getDataRange().getValues();
    for (var i = 1; i < dataKp.length; i++) {
      var dbTahun = dataKp[i][1] ? dataKp[i][1].toString().trim() : "";
      var dbCadangan = dataKp[i][11] ? dataKp[i][11].toString().trim().toUpperCase() : ""; 
      
      if (dbTahun === tahunSemasa && dbCadangan !== "") {
        kiraanCadangan[dbCadangan] = (kiraanCadangan[dbCadangan] || 0) + 1;
      }
    }
  }
  
  var senaraiSekolah = [];
  var isSekAsalSKM_Pedalaman = false;
  var kategoriSekAsal = ""; 

  for (var j = 1; j < data.length; j++) {
    var dbKodA = data[j][5] ? data[j][5].toString().trim().toUpperCase() : "";
    if (dbKodA === kodSekAsal.toUpperCase()) {
      var bilCalonAsal = data[j][23] ? parseInt(data[j][23]) : 0;
      var statusLokasiAsal = data[j][26] ? data[j][26].toString().trim().toUpperCase() : ""; // TANGKAP LOKASI

      if ((bilCalonAsal > 0 && bilCalonAsal < 20) || statusLokasiAsal === "PEDALAMAN") {
         isSekAsalSKM_Pedalaman = true;
      }
      kategoriSekAsal = data[j][1] ? data[j][1].toString().trim().toUpperCase() : ""; 
      break;
    }
  }

  for (var i = 1; i < data.length; i++) {
    var dbPpd = data[i][3] ? data[i][3].toString().trim().toUpperCase() : ""; 
    var dbKod = data[i][5] ? data[i][5].toString().trim().toUpperCase() : "";
    var dbNama = data[i][6] ? data[i][6].toString().trim().toUpperCase() : "";
    var dbBilCalon = data[i][23] ? parseInt(data[i][23]) : 0;
    var dbKategori = data[i][1] ? data[i][1].toString().trim().toUpperCase() : "";
    var dbLokasi = data[i][26] ? data[i][26].toString().trim().toUpperCase() : ""; 

    var teksTagging = "";
    if (dbLokasi === "PEDALAMAN") {
        teksTagging = " [PEDALAMAN]";
    } else if (dbBilCalon > 0 && dbBilCalon < 20) {
        teksTagging = " [SKM]";
    }

    if (dbPpd === kodPpdSemasa.toUpperCase() && dbKategori === kategoriSekAsal) {
      if (!isSekAsalSKM_Pedalaman && dbKod === kodSekAsal.toUpperCase()) continue;

      var jumlahSemasa = kiraanCadangan[dbKod] || 0;
      var teksKuota = jumlahSemasa >= 1 ? " [1/1 - PENUH]" : " [0/1 - KOSONG]";
      
      senaraiSekolah.push({ 
        kod: dbKod, 
        nama: dbKod + " - " + dbNama + teksTagging + teksKuota 
      });
    }
  }
  return senaraiSekolah;
}

function simpanCadanganKP(d) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ambilSheetKalisRalat(ss, 'tbkp');
  var data = sheet.getDataRange().getValues();
  
  var inputKp = d.no_kp.toString().trim();
  var tahunSemasa = d.tahun.toString().trim();
  var tahapSemasa = d.matrik_tahap.toString().trim();
  var idMatrikTahap = (tahapSemasa === "TAHUN 4") ? "MP4" : "MP3";
  var idUnikSurat = tahunSemasa + d.id_jawatan + idMatrikTahap;
  var rEdit = d.rowNum ? Number(d.rowNum) : -1;
  
  for (var i = 1; i < data.length; i++) {
    var barisSemasaDB = i + 1;
    if (barisSemasaDB === rEdit) continue; 
    
    var dbTahun = data[i][1] ? data[i][1].toString().trim() : "";
    var dbTahap = data[i][4] ? data[i][4].toString().trim() : "";
    var dbKp = data[i][7] ? data[i][7].toString().trim() : "";
    
    if (dbKp === inputKp && dbTahun === tahunSemasa && dbTahap === tahapSemasa) {
      return { status: "ERROR", message: "Kad Pengenalan ini telah dicalonkan untuk " + tahapSemasa + " tahun " + tahunSemasa };
    }
  }
  
  if (rEdit > 0) {
    sheet.getRange(rEdit, 5).setValue(tahapSemasa);
    sheet.getRange(rEdit, 7).setValue(d.nama_guru.toUpperCase()); 
    sheet.getRange(rEdit, 8).setValue(inputKp); 
    sheet.getRange(rEdit, 9).setValue(d.no_tel); 
    sheet.getRange(rEdit, 10).setValue(d.emel); 
    sheet.getRange(rEdit, 11).setValue(d.alamat_rumah.toUpperCase()); 
    sheet.getRange(rEdit, 12).setValue(d.kod_sek_cadangan); 
    sheet.getRange(rEdit, 15).setValue(idUnikSurat); 
    
    SpreadsheetApp.flush();
    return { status: "SUCCESS", message: "Cadangan Ketua Pengawas berjaya dikemaskini!" };
  } else {
    var nextId = data.length > 1 ? Number(data[data.length - 1][0]) + 1 : 1;
    sheet.appendRow([
      nextId, tahunSemasa, d.kod_sek_asal, d.daerah, tahapSemasa, d.id_jawatan, 
      d.nama_guru.toUpperCase(), inputKp, d.no_tel, d.emel, d.alamat_rumah.toUpperCase(), 
      d.kod_sek_cadangan, "", "CADANGAN", idUnikSurat
    ]);
    return { status: "SUCCESS", message: "Cadangan Ketua Pengawas berjaya dihantar!" };
  }
}

function dapatkanSenaraiKPBySekolah(kodSekAsal, tahunSemasa) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ambilSheetKalisRalat(ss, 'tbkp');
  var data = sheet.getDataRange().getValues();
  
  var sheetSek = ss.getSheetByName('tbsekolah');
  var dataSek = sheetSek ? sheetSek.getDataRange().getValues() : [];
  var mapSekolah = {};
  for(var s = 1; s < dataSek.length; s++) {
    var kodS = dataSek[s][5] ? dataSek[s][5].toString().trim().toUpperCase() : "";
    var namaS = dataSek[s][6] ? dataSek[s][6].toString().trim().toUpperCase() : "";
    if (kodS !== "") {
      mapSekolah[kodS] = namaS;
    }
  }
  
  var list = [];
  for (var i = 1; i < data.length; i++) {
    var dbTahun = data[i][1] ? data[i][1].toString().trim() : "";
    var dbSekAsal = data[i][2] ? data[i][2].toString().trim() : "";
    if (dbSekAsal === kodSekAsal && dbTahun === tahunSemasa.toString()) {
      var kodCadang = data[i][11] ? data[i][11].toString().trim().toUpperCase() : "";
      var namaCadang = mapSekolah[kodCadang] ? mapSekolah[kodCadang] : "";
      var paparanCadang = namaCadang ? (kodCadang + " - " + namaCadang) : kodCadang; 

      list.push({
        rowNum: i + 1, id: data[i][0], matrik_tahap: data[i][4], nama_guru: data[i][6],
        no_kp: data[i][7], no_tel: data[i][8], emel: data[i][9], alamat_rumah: data[i][10], 
        kod_cadangan: kodCadang, paparan_cadangan: paparanCadang, status: data[i][13]
      });
    }
  }
  return list;
}

// =========================================================================
// 9. MODUL PPD (SIMPAN PENETAPAN OPERASI - BK, PK, KP)
// =========================================================================
function simpanDataOperasiPPD(d) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetMatrik = ambilSheetKalisRalat(ss, 'tbmatrik_operasi');
    var dataMatrik = sheetMatrik.getDataRange().getValues();
    var barisKemaskini = -1;
    
    for (var i = 1; i < dataMatrik.length; i++) {
      var dbTahun = dataMatrik[i][1] ? dataMatrik[i][1].toString().trim() : "";
      var dbKodSek = dataMatrik[i][2] ? dataMatrik[i][2].toString().trim().toUpperCase() : "";
      var dbTahap = dataMatrik[i][4] ? dataMatrik[i][4].toString().trim().toUpperCase() : ""; 
      if (dbTahun === d.tahun && dbKodSek === d.kod_sekolah && dbTahap === d.tahap) {
        barisKemaskini = i + 1;
        break;
      }
    }
    
    if (barisKemaskini !== -1) {
      sheetMatrik.getRange(barisKemaskini, 6).setValue(d.id_bilik_kebal);
      sheetMatrik.getRange(barisKemaskini, 7).setValue(d.no_kp_pk1);
      sheetMatrik.getRange(barisKemaskini, 8).setValue(d.no_kp_pk2);
    } else {
      var nextId = dataMatrik.length > 1 ? Number(dataMatrik[dataMatrik.length - 1][0]) + 1 : 1;
      sheetMatrik.appendRow([
        nextId, d.tahun, d.kod_sekolah, d.daerah, d.tahap, d.id_bilik_kebal, d.no_kp_pk1, d.no_kp_pk2
      ]);
    }

    var sheetKp = ss.getSheetByName('tbkp');
    if (sheetKp) {
      var dataKp = sheetKp.getDataRange().getValues();
      for (var k = 1; k < dataKp.length; k++) {
        var dbTahunKp = dataKp[k][1] ? dataKp[k][1].toString().trim() : "";
        var dbTahapKp = dataKp[k][4] ? dataKp[k][4].toString().trim().toUpperCase() : "";
        var dbSekTetap = dataKp[k][12] ? dataKp[k][12].toString().trim().toUpperCase() : "";
        
        if (dbTahunKp === d.tahun && dbSekTetap === d.kod_sekolah && dbTahapKp === d.tahap) {
          sheetKp.getRange(k + 1, 13).setValue(""); 
          sheetKp.getRange(k + 1, 14).setValue("CADANGAN"); 
        }
      }

      if (d.no_kp_kp !== "") {
        for (var k = 1; k < dataKp.length; k++) {
          var dbTahunKp = dataKp[k][1] ? dataKp[k][1].toString().trim() : "";
          var dbNoKp = dataKp[k][7] ? dataKp[k][7].toString().trim() : "";
          var dbTahapKp = dataKp[k][4] ? dataKp[k][4].toString().trim().toUpperCase() : "";
          
          if (dbTahunKp === d.tahun && dbNoKp === d.no_kp_kp && dbTahapKp === d.tahap) {
             sheetKp.getRange(k + 1, 13).setValue(d.kod_sekolah); 
             sheetKp.getRange(k + 1, 14).setValue(d.status_kp); 
             break;
          }
        }
      }
    }
    SpreadsheetApp.flush();
    return { status: "SUCCESS" };
  } catch (e) { return { status: "FAILED", message: e.message }; }
}

// =========================================================================
// SIMPAN / KEMASKINI PENYELIA KAWASAN (MODAL UI UNGU) - BERKUNCI BERKEMBAR
// =========================================================================
function simpanDataPkBackend(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // KUNCI: Elak ID bertindih
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ambilSheetKalisRalat(ss, 'tbpk');
    var data = sheet.getDataRange().getValues();
    var r = payload.rowNum ? Number(payload.rowNum) : 0;
    
    var kodDaerahSebenar = "PPD";
    try {
      var sheetSek = ss.getSheetByName('tbsekolah');
      var dataSek = sheetSek ? sheetSek.getDataRange().getValues() : [];
      for (var s = 1; s < dataSek.length; s++) {
        var kodS = dataSek[s][5] ? dataSek[s][5].toString().trim().toUpperCase() : "";
        if (kodS === payload.kod_sek_asal.trim().toUpperCase()) {
          var ppdSek = dataSek[s][3] ? dataSek[s][3].toString().trim().toUpperCase() : "";
          if (ppdSek !== "") kodDaerahSebenar = ppdSek;
          break;
        }
      }
    } catch(errDae) {}

    // PENGAWAL KESELAMATAN BERKUNCI 3 SERANGKAI (Tahun + Tahap + IC)
    var inputTahun = payload.tahun ? payload.tahun.toString().trim() : "";
    var inputTahap = payload.matrik_tahap ? payload.matrik_tahap.toString().trim().toUpperCase() : "";
    var inputKp = payload.no_kp ? payload.no_kp.toString().trim() : "";

    for (var i = 1; i < data.length; i++) {
      if (r > 0 && r === (i + 1)) continue;
      
      var dbTahun = data[i][1] ? data[i][1].toString().trim() : "";
      var dbTahap = data[i][4] ? data[i][4].toString().trim().toUpperCase() : "";
      var dbKp = data[i][8] ? data[i][8].toString().trim() : "";
      
      if (dbTahun === inputTahun && dbTahap === inputTahap && dbKp === inputKp) {
        lock.releaseLock();
        return { status: "WARNING", message: "Kad Pengenalan ini telah didaftarkan sebagai Penyelia Kawasan untuk tahap (" + inputTahap + ") pada sesi ini!" };
      }
    }

    var idMatrikTahap = (inputTahap === "TAHUN 4") ? "MP4" : "MP3";
    var idUnikSurat = inputTahun + "PK01" + idMatrikTahap;

    if (r > 1 && r <= data.length) {
      sheet.getRange(r, 2).setValue(inputTahun);
      sheet.getRange(r, 3).setValue(payload.kod_sek_asal);
      sheet.getRange(r, 4).setValue(kodDaerahSebenar);
      sheet.getRange(r, 5).setValue(inputTahap);
      sheet.getRange(r, 6).setValue("PK01");
      sheet.getRange(r, 7).setValue(payload.jawatan_pk);
      sheet.getRange(r, 8).setValue(payload.nama_pk.toUpperCase());
      sheet.getRange(r, 9).setValue(inputKp);
      sheet.getRange(r, 10).setValue(payload.no_tel);
      sheet.getRange(r, 11).setValue(payload.emel);
      sheet.getRange(r, 12).setValue(payload.alamat_rumah ? payload.alamat_rumah.toUpperCase() : "");
      sheet.getRange(r, 13).setValue(idUnikSurat);
    } else {
      var maxId = 0;
      for (var k = 1; k < data.length; k++) {
        var idVal = parseInt(data[k][0]) || 0;
        if (idVal > maxId) maxId = idVal;
      }
      var nextId = maxId + 1; // ID Terkini secara live
      
      sheet.appendRow([
        nextId, inputTahun, payload.kod_sek_asal, kodDaerahSebenar, inputTahap, "PK01",
        payload.jawatan_pk, payload.nama_pk.toUpperCase(), inputKp, payload.no_tel, payload.emel,
        payload.alamat_rumah ? payload.alamat_rumah.toUpperCase() : "", idUnikSurat
      ]);
    }
    SpreadsheetApp.flush();
    lock.releaseLock();
    return { status: "SUCCESS", message: "Maklumat Penyelia Kawasan berjaya disimpan!" };
  } catch (e) {
    if (lock) lock.releaseLock();
    return { status: "FAILED", message: "Ralat Server: " + e.message };
  }
}

// =========================================================================
// 10. MODUL PENDAFTARAN PENYELIA KAWASAN (ROLE: PPD/JPN) - BERKUNCI BERKEMBAR
// =========================================================================
function urusDataPkPPD(kodDaerah, d, rowNum) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ambilSheetKalisRalat(ss, 'tbpk');
    var r = rowNum ? Number(rowNum) : 0;
    
    var kodDaerahSebenar = kodDaerah.toUpperCase();
    try {
      var sheetSek = ss.getSheetByName('tbsekolah');
      var dataSek = sheetSek ? sheetSek.getDataRange().getValues() : [];
      for (var s = 1; s < dataSek.length; s++) {
        var kodS = dataSek[s][5] ? dataSek[s][5].toString().trim().toUpperCase() : "";
        if (kodS === d.kod_sek_asal.trim().toUpperCase()) {
          var ppdSek = dataSek[s][3] ? dataSek[s][3].toString().trim().toUpperCase() : "";
          if (ppdSek !== "") kodDaerahSebenar = ppdSek; 
          break;
        }
      }
    } catch(errDae) {}

    // PENGAWAL KESELAMATAN BERKUNCI 3 SERANGKAI (Tahun + Tahap + IC)
    var inputTahun = d.tahun ? d.tahun.toString().trim() : "";
    var inputTahap = d.matrik_tahap ? d.matrik_tahap.toString().trim().toUpperCase() : "";
    var inputKp = d.no_kp ? d.no_kp.toString().trim() : "";

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (r > 0 && r === (i + 1)) continue; 
      
      var dbTahun = data[i][1] ? data[i][1].toString().trim() : "";
      var dbTahap = data[i][4] ? data[i][4].toString().trim().toUpperCase() : "";
      var dbKp = data[i][8] ? data[i][8].toString().trim() : ""; 
      
      if (dbTahun === inputTahun && dbTahap === inputTahap && dbKp === inputKp) {
        lock.releaseLock();
        return { status: "FAILED", message: "Kad Pengenalan ini telah didaftarkan sebagai Penyelia Kawasan untuk tahap (" + inputTahap + ") pada sesi ini!" };
      }
    }

    var idMatrikTahap = (inputTahap === "TAHUN 4") ? "MP4" : "MP3";
    var idUnikSurat = inputTahun + "PK01" + idMatrikTahap; 

    if (r > 0) {
      sheet.getRange(r, 3).setValue(d.kod_sek_asal);
      sheet.getRange(r, 4).setValue(kodDaerahSebenar); 
      sheet.getRange(r, 5).setValue(inputTahap);
      sheet.getRange(r, 6).setValue("PK01");
      sheet.getRange(r, 7).setValue(d.jawatan_pk);
      sheet.getRange(r, 8).setValue(d.nama_pk.toUpperCase());
      sheet.getRange(r, 9).setValue(inputKp);
      sheet.getRange(r, 10).setValue(d.no_tel);
      sheet.getRange(r, 11).setValue(d.emel);
      sheet.getRange(r, 12).setValue(d.alamat_rumah ? d.alamat_rumah.toUpperCase() : "");
      sheet.getRange(r, 13).setValue(idUnikSurat);
    } else {
      var maxId = 0;
      for (var k = 1; k < data.length; k++) {
        var idVal = parseInt(data[k][0]) || 0;
        if (idVal > maxId) maxId = idVal;
      }
      var nextId = maxId + 1; // ID Terkini secara live
      
      sheet.appendRow([
        nextId, inputTahun, d.kod_sek_asal, kodDaerahSebenar, inputTahap, "PK01", 
        d.jawatan_pk, d.nama_pk.toUpperCase(), inputKp, d.no_tel, d.emel,
        d.alamat_rumah ? d.alamat_rumah.toUpperCase() : "", idUnikSurat
      ]);
    }
    SpreadsheetApp.flush();
    lock.releaseLock();
    return { status: "SUCCESS" };
  } catch (e) { 
    if (lock) lock.releaseLock();
    return { status: "FAILED", message: e.message }; 
  }
}

// =========================================================================
// 11. FUNGSI PENGESAHAN STATUS SEKOLAH
// =========================================================================
function kemaskiniStatusPengesahanSekolah(rowNum, statusDipilih) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ambilSheetKalisRalat(ss, 'tbsekolah');
    sheet.getRange(Number(rowNum), 25).setValue(statusDipilih.toUpperCase());
    SpreadsheetApp.flush();
    return "SUCCESS";
  } catch (e) { throw new Error(e.message); }
}

// =========================================================================
// 12. FUNGSI BULK CLONE: SALIN PK KE TAHUN HADAPAN
// =========================================================================
function salinDataPkKeTahun(kodDaerah, senaraiRowNum, tahunSasar) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('tbpk');
    var data = sheet.getDataRange().getValues();
    var header = data[0];
    
    var idxTahun = header.indexOf('tahun');
    var idxKp = header.indexOf('no_kp');
    var idxTahap = header.indexOf('matrik_tahap');
    var idxDaerah = header.indexOf('daerah');
    
    var pkWujud = {};
    for (var i = 1; i < data.length; i++) {
      if (data[i][idxTahun].toString() === tahunSasar.toString()) {
        var kunci = data[i][idxDaerah] + "_" + data[i][idxKp] + "_" + data[i][idxTahap];
        pkWujud[kunci] = true;
      }
    }
    
    var rekodBaru = [];
    var jumlahDisalin = 0;
    var jumlahPendua = 0;
    
    senaraiRowNum.forEach(function(rowNum) {
      var rowIdx = parseInt(rowNum) - 1; 
      if (data[rowIdx]) {
        var barisAsal = data[rowIdx].slice(); 
        var kunciSemak = barisAsal[idxDaerah] + "_" + barisAsal[idxKp] + "_" + barisAsal[idxTahap];
        
        if (!pkWujud[kunciSemak]) {
          barisAsal[idxTahun] = tahunSasar; 
          rekodBaru.push(barisAsal);
          pkWujud[kunciSemak] = true; 
          jumlahDisalin++;
        } else { jumlahPendua++; }
      }
    });
    
    if (rekodBaru.length > 0) {
      var startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, rekodBaru.length, rekodBaru[0].length).setValues(rekodBaru);
    }
    
    return {
      status: "SUCCESS",
      message: "Berjaya menyalin " + jumlahDisalin + " rekod PK ke sesi " + tahunSasar + "! (" + jumlahPendua + " rekod diabaikan kerana telah wujud/pendua)."
    };
  } catch(e) { return { status: "ERROR", message: "Ralat Backend: " + e.toString() }; }
}

// =========================================================================
// FUNGSI SALIN PUKAL PK (SEKATAN PENDUA & TAHUN SAMA - TURBO)
// =========================================================================
function salinPukalPkBackend(senaraiId, tahunSasaran, tahap) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ambilSheetKalisRalat(ss, 'tbpk');
    var data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return { status: "FAILED", message: "Tiada data Penyelia Kawasan untuk disalin." };
    }

    var header = data[0].map(function(h) { return h.toString().trim().toLowerCase(); });
    var idxId = header.indexOf('id') !== -1 ? header.indexOf('id') : 0;
    var idxTahun = header.indexOf('tahun') !== -1 ? header.indexOf('tahun') : 1;
    var idxDaerah = header.indexOf('daerah') !== -1 ? header.indexOf('daerah') : 3;
    var idxTahap = header.indexOf('matrik_tahap') !== -1 ? header.indexOf('matrik_tahap') : 4;
    var idxIdJawatan = header.indexOf('id_jawatan') !== -1 ? header.indexOf('id_jawatan') : 5;
    var idxKp = header.indexOf('no_kp') !== -1 ? header.indexOf('no_kp') : 8;
    var idxRujukanSurat = header.indexOf('no_rujukan_surat') !== -1 ? header.indexOf('no_rujukan_surat') : 12;

    var tahapSurat = (tahap === 'RENDAH') ? 'TAHUN 4' : 'TINGKATAN 3';

    var maxId = 0;
    for (var i = 1; i < data.length; i++) {
      var currId = Number(data[i][idxId]);
      if (!isNaN(currId) && currId > maxId) maxId = currId;
    }

    var pkWujud = {};
    for (var i = 1; i < data.length; i++) {
      if (data[i][idxTahun] && data[i][idxTahun].toString().trim() === tahunSasaran.toString().trim()) {
        var kunci = data[i][idxDaerah] + "_" + data[i][idxKp] + "_" + data[i][idxTahap];
        pkWujud[kunci] = true;
      }
    }

    var rekodBaru = [];
    var jumlahDisalin = 0;
    var jumlahPendua = 0;

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowTahap = row[idxTahap] ? row[idxTahap].toString().trim().toUpperCase() : "";
      var rowKp = row[idxKp] ? row[idxKp].toString().trim() : "";

      if (rowTahap === tahapSurat.toUpperCase()) {
        if (!senaraiId || senaraiId.length === 0 || senaraiId.indexOf(rowKp) !== -1) {
          var barisAsal = row.slice();
          var kunciSemak = barisAsal[idxDaerah] + "_" + barisAsal[idxKp] + "_" + barisAsal[idxTahap];

          if (!pkWujud[kunciSemak]) {
            maxId++;
            if (idxId !== -1) barisAsal[idxId] = maxId;
            if (idxTahun !== -1) barisAsal[idxTahun] = tahunSasaran;

            if (idxRujukanSurat !== -1) {
              var idMatrikTahap = (rowTahap === "TAHUN 4") ? "MP4" : "MP3";
              var idJawatan = (idxIdJawatan !== -1 && barisAsal[idxIdJawatan]) ? barisAsal[idxIdJawatan].toString().trim() : "PK01";
              barisAsal[idxRujukanSurat] = tahunSasaran + idJawatan + idMatrikTahap;
            }

            rekodBaru.push(barisAsal);
            pkWujud[kunciSemak] = true;
            jumlahDisalin++;
          } else {
            jumlahPendua++;
          }
        }
      }
    }

    if (rekodBaru.length === 0) {
      return {
        status: "WARNING",
        message: "⚠️ Operasi Ditolak: Semua rekod PK yang dipilih telah pun wujud dalam Sesi " + tahunSasaran + "!"
      };
    }

    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rekodBaru.length, rekodBaru[0].length).setValues(rekodBaru);

    return {
      status: "SUCCESS",
      message: "Berjaya menyalin " + jumlahDisalin + " rekod PK ke Sesi " + tahunSasaran + "! (" + jumlahPendua + " rekod pendua diabaikan)."
    };

  } catch (e) {
    return { status: "FAILED", message: "Ralat Backend: " + e.toString() };
  }
}

// =========================================================================
// 13. MODUL PENGURUSAN KETUA PENGAWAS (ROLE: PPD/JPN) - KEMASKINI PEDALAMAN
// =========================================================================
function urusDataKpPPD(kodDaerah, d, rowNum) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
  } catch (e) {
    return { status: "FAILED", message: "⏳ Sistem sedang memproses data pengguna lain. Sila klik simpan semula dalam 2 saat." };
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Semak Syarat SOP (SKM & Sekolah Asal)
    var isSKM_Pedalaman = false;
    try {
      var sheetSek = ss.getSheetByName('tbsekolah');
      var dataSek = sheetSek.getDataRange().getValues();
      for(var s = 1; s < dataSek.length; s++) {
        var k = dataSek[s][5] ? dataSek[s][5].toString().trim().toUpperCase() : "";
        var bilCalon = dataSek[s][23] ? parseInt(dataSek[s][23]) : 0;
        var statusLokasi = dataSek[s][26] ? dataSek[s][26].toString().trim().toUpperCase() : "";
        
        if (k === d.kod_sek_asal.trim().toUpperCase()) {
           if ((bilCalon > 0 && bilCalon < 20) || statusLokasi === "PEDALAMAN") {
              isSKM_Pedalaman = true; 
           }
           break;
        }
      }
    } catch(e){}

    if (!isSKM_Pedalaman && d.kod_sek_tetap && d.kod_sek_tetap.trim() !== "" && d.kod_sek_tetap.trim().toUpperCase() === d.kod_sek_asal.trim().toUpperCase()) {
      lock.releaseLock();
      return { status: "FAILED", message: "⛔ Ralat SOP: Ketua Pengawas TIDAK BOLEH ditempatkan di sekolah asal mereka sendiri melainkan SKM atau Sekolah Pedalaman!" };
    }

    var sheet = ambilSheetKalisRalat(ss, 'tbkp');
    var data = sheet.getDataRange().getValues();
    var r = rowNum ? Number(rowNum) : 0;
    var realRow = 0;

    // 2. LOGIK KALIS DELETE (ROW SHIFT HUNTING)
    if (r > 0) {
       var oldKp = data[r - 1] ? (data[r - 1][7] ? data[r - 1][7].toString().trim() : "") : "";
       
       if (oldKp === d.no_kp.toString().trim()) {
           realRow = r; // Data tidak berganjak, selamat untuk ditulis
       } else {
           // Data telah berganjak! Sistem memburu baris sebenar secara live:
           for (var j = 1; j < data.length; j++) {
               var dbTahunCari = data[j][1] ? data[j][1].toString().trim() : "";
               var dbTahapCari = data[j][4] ? data[j][4].toString().trim().toUpperCase() : "";
               var dbKpCari = data[j][7] ? data[j][7].toString().trim() : "";
               
               if (dbTahunCari === d.tahun.toString().trim() && 
                   dbTahapCari === d.matrik_tahap.toString().trim().toUpperCase() && 
                   dbKpCari === d.no_kp.toString().trim()) {
                   realRow = j + 1;
                   break;
               }
           }
       }
       
       // Jika masih tak jumpa, bermakna rekod tu dah dibuang terus. Hentikan proses!
       if (realRow === 0) {
           lock.releaseLock();
           return { status: "FAILED", message: "Ralat Sinkronisasi: Rekod ini telah berganjak atau dipadam sepenuhnya. Sila 'Refresh' halaman anda." };
       }
    }

    // 3. Dapatkan Kod Daerah Sebenar
    var kodDaerahSebenar = kodDaerah.toUpperCase();
    try {
      var sheetSekolah = ss.getSheetByName('tbsekolah');
      var dataSekolah = sheetSekolah ? sheetSekolah.getDataRange().getValues() : [];
      for (var s = 1; s < dataSekolah.length; s++) {
        var kodS = dataSekolah[s][5] ? dataSekolah[s][5].toString().trim().toUpperCase() : "";
        if (kodS === d.kod_sek_asal.trim().toUpperCase()) {
          var ppdSek = dataSekolah[s][3] ? dataSekolah[s][3].toString().trim().toUpperCase() : ""; 
          if (ppdSek !== "") kodDaerahSebenar = ppdSek; 
          break;
        }
      }
    } catch(errDae) {}

    // 4. Pengawal Keselamatan Pendua (Tahun + Tahap + IC)
    var inputTahun = d.tahun ? d.tahun.toString().trim() : "";
    var inputTahap = d.matrik_tahap ? d.matrik_tahap.toString().trim().toUpperCase() : "";
    var inputKp = d.no_kp ? d.no_kp.toString().trim() : "";

    for (var i = 1; i < data.length; i++) {
      if (realRow > 0 && realRow === (i + 1)) continue;
      
      var dbTahun = data[i][1] ? data[i][1].toString().trim() : "";
      var dbTahap = data[i][4] ? data[i][4].toString().trim().toUpperCase() : "";
      var dbKp = data[i][7] ? data[i][7].toString().trim() : "";
      
      if (dbTahun === inputTahun && dbTahap === inputTahap && dbKp === inputKp) {
        lock.releaseLock();
        return { status: "FAILED", message: "Kad Pengenalan ini telah didaftarkan sebagai Ketua Pengawas untuk tahap & sesi ini!" };
      }
    }

    // 5. Simpan Data di Baris Yang Betul
    var idMatrikTahap = (inputTahap === "TAHUN 4") ? "MP4" : "MP3";
    var idUnikSurat = inputTahun + "KP02" + idMatrikTahap;

    if (realRow > 0) {
      sheet.getRange(realRow, 3).setValue(d.kod_sek_asal);
      sheet.getRange(realRow, 4).setValue(kodDaerahSebenar); 
      sheet.getRange(realRow, 5).setValue(inputTahap);
      sheet.getRange(realRow, 6).setValue("KP02");
      sheet.getRange(realRow, 7).setValue(d.nama_guru.toUpperCase());
      sheet.getRange(realRow, 8).setValue(inputKp);
      sheet.getRange(realRow, 9).setValue(d.no_tel);
      sheet.getRange(realRow, 10).setValue(d.emel);
      sheet.getRange(realRow, 11).setValue(d.alamat_rumah ? d.alamat_rumah.toUpperCase() : "");
      sheet.getRange(realRow, 12).setValue(d.kod_sek_cadangan ? d.kod_sek_cadangan.toUpperCase() : "");
      sheet.getRange(realRow, 13).setValue(d.kod_sek_tetap ? d.kod_sek_tetap.toUpperCase() : "");
      sheet.getRange(realRow, 14).setValue(d.status_penetapan ? d.status_penetapan.toUpperCase() : "CADANGAN");
      sheet.getRange(realRow, 15).setValue(idUnikSurat);
    } else {
      var maxId = 0;
      for (var k = 1; k < data.length; k++) {
        var idVal = parseInt(data[k][0]) || 0;
        if (idVal > maxId) maxId = idVal;
      }
      var nextId = maxId + 1; 
      
      sheet.appendRow([
        nextId, inputTahun, d.kod_sek_asal, kodDaerahSebenar, inputTahap, "KP02", 
        d.nama_guru.toUpperCase(), inputKp, d.no_tel, d.emel,
        d.alamat_rumah ? d.alamat_rumah.toUpperCase() : "",
        d.kod_sek_cadangan ? d.kod_sek_cadangan.toUpperCase() : "",
        d.kod_sek_tetap ? d.kod_sek_tetap.toUpperCase() : "",
        d.status_penetapan ? d.status_penetapan.toUpperCase() : "CADANGAN", idUnikSurat
      ]);
    }
    SpreadsheetApp.flush();
    lock.releaseLock();
    return { status: "SUCCESS" };
  } catch (e) {
    if (lock) lock.releaseLock();
    return { status: "FAILED", message: e.message };
  } 
}

// =========================================================================
// SUIS KAWALAN KEBENARAN CETAK KP (ADMIN & JPN) - PENELATAN TAHAP
// =========================================================================
function normalisasikanTahap(tahapInput) {
  if (!tahapInput) return "RENDAH";
  var t = tahapInput.toString().toUpperCase().trim();
  if (t.indexOf("4") !== -1 || t === "RENDAH") return "RENDAH";
  if (t.indexOf("3") !== -1 || t === "MENENGAH") return "MENENGAH";
  return t;
}

function dapatkanStatusKebenaranCetakKP(tahun, tahap) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("tb_tetapan"); 
    if (!sheet) return { status: "SUCCESS", kebenaranCetak: "TIDAK" };
    
    var tahapNorm = normalisasikanTahap(tahap);
    var data = sheet.getDataRange().getValues();
    var kunciCari = "KP_CETAK_" + tahun + "_" + tahapNorm;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === kunciCari) {
        return { status: "SUCCESS", kebenaranCetak: data[i][1] || "TIDAK" };
      }
    }
    return { status: "SUCCESS", kebenaranCetak: "TIDAK" };
  } catch (e) {
    return { status: "ERROR", message: e.message };
  }
}

function simpanStatusKebenaranCetakKP(tahun, tahap, statusBaharu) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("tb_tetapan");
    if (!sheet) {
      sheet = ss.insertSheet("tb_tetapan");
      sheet.appendRow(["kunci", "nilai"]);
    }
    
    var tahapNorm = normalisasikanTahap(tahap);
    var data = sheet.getDataRange().getValues();
    var kunciCari = "KP_CETAK_" + tahun + "_" + tahapNorm;
    var wujud = false;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === kunciCari) {
        sheet.getRange(i + 1, 2).setValue(statusBaharu);
        wujud = true;
        break;
      }
    }
    
    if (!wujud) {
      sheet.appendRow([kunciCari, statusBaharu]);
    }
    
    return { status: "SUCCESS" };
  } catch (e) {
    return { status: "ERROR", message: e.message };
  }
}

// =========================================================================
// FUNGSI TURBO: SIMPAN PENEMPATAN INLINE PPD (KALIS ROW-SHIFT)
// =========================================================================
function simpanPenempatanInlineKp(rowNum, kodSekTetap, statusPenetapan, kodSekAsal) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
  } catch (e) {
    return { status: "FAILED", message: "⏳ Sistem sedang memproses data pengguna lain. Sila klik simpan semula dalam 2 saat." };
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Semakan SOP (SKM/PEDALAMAN)
    var isSKM_Pedalaman = false;
    try {
      var sheetSek = ss.getSheetByName('tbsekolah');
      var dataSek = sheetSek.getDataRange().getValues();
      for(var s=1; s<dataSek.length; s++) {
        var k = dataSek[s][5] ? dataSek[s][5].toString().trim().toUpperCase() : "";
        var bilCalon = dataSek[s][23] ? parseInt(dataSek[s][23]) : 0;
        var statusLokasi = dataSek[s][26] ? dataSek[s][26].toString().trim().toUpperCase() : "";
        
        if(k === kodSekAsal.trim().toUpperCase()) {
           if ((bilCalon > 0 && bilCalon < 20) || statusLokasi === "PEDALAMAN") {
               isSKM_Pedalaman = true; 
           }
           break;
        }
      }
    } catch(e){}

    if (!isSKM_Pedalaman && kodSekTetap && kodSekTetap.trim() !== "" && kodSekTetap.trim().toUpperCase() === kodSekAsal.trim().toUpperCase()) {
      lock.releaseLock();
      return { status: "FAILED", message: "⛔ Ralat SOP: Ketua Pengawas TIDAK BOLEH ditempatkan di sekolah asal mereka sendiri (" + kodSekAsal + ") melainkan SKM/PEDALAMAN!" };
    }

    var sheet = ambilSheetKalisRalat(ss, 'tbkp');
    var data = sheet.getDataRange().getValues();
    var r = Number(rowNum);

    if (r <= 1 || r > data.length) {
      lock.releaseLock();
      return { status: "FAILED", message: "Nombor baris tidak sah! Sila 'Refresh' halaman." };
    }

    // 2. PENGAWAL DATA BASI (Stale Data Checker)
    var dbSekAsal = data[r-1][2] ? data[r-1][2].toString().trim().toUpperCase() : "";
    var expectedSekAsal = kodSekAsal ? kodSekAsal.trim().toUpperCase() : "";
    
    if (expectedSekAsal !== "" && dbSekAsal !== expectedSekAsal) {
        lock.releaseLock();
        return { status: "FAILED", message: "Ralat Sinkronisasi: Baris jadual telah berganjak akibat pemadaman rekod oleh pengguna lain. Sila 'Refresh' skrin anda sebelum membuat pengesahan." };
    }

    if (expectedSekAsal === "") {
        var dbStatus = data[r-1][13] ? data[r-1][13].toString().trim().toUpperCase() : "";
        if (dbStatus !== "DISAHKAN") {
            lock.releaseLock();
            return { status: "FAILED", message: "Ralat Sinkronisasi: Data telah berganjak. Sila 'Refresh' skrin anda." };
        }
    }

    // 3. Lulus & Tulis Penempatan
    sheet.getRange(r, 13).setValue(kodSekTetap.toUpperCase().trim());
    sheet.getRange(r, 14).setValue(statusPenetapan.toUpperCase().trim());

    SpreadsheetApp.flush();
    lock.releaseLock();
    return { status: "SUCCESS", message: "Penempatan berjaya disahkan!" };
  } catch (e) {
    if (lock) lock.releaseLock();
    return { status: "FAILED", message: e.message };
  }
}

// =========================================================================
// MODUL KETUA PENTAKSIR KPN / KPKw (KEMASKINI DENGAN LIVE ID & LOCK)
// =========================================================================
function urusDataKpnKpkwPPD(d, rowNum) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // KUNCI: Elakkan PPD bertembung masa simpan data
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ambilSheetKalisRalat(ss, 'tbkpn_kpkw');
    var r = rowNum ? Number(rowNum) : 0;
    
    var mapSubjek = {};
    try {
      var sheetMp = ss.getSheetByName('tbmp');
      if (sheetMp) {
        var dataMp = sheetMp.getDataRange().getValues();
        for (var m = 1; m < dataMp.length; m++) {
          var kodM = dataMp[m][1] ? dataMp[m][1].toString().trim().toUpperCase() : "";
          var namaM = dataMp[m][2] ? dataMp[m][2].toString().trim().toUpperCase() : "";
          if (kodM !== "") mapSubjek[kodM] = namaM;
        }
      }
    } catch(eMp) {}

    var data = sheet.getDataRange().getValues();
    
    var inputTahun = d.tahun ? d.tahun.toString().trim() : "";
    var inputTahap = d.matrik_tahap ? d.matrik_tahap.toString().trim().toUpperCase() : "";
    var inputPeringkat = d.peringkat ? d.peringkat.toString().trim().toUpperCase() : "";
    var inputDaerah = d.daerah ? d.daerah.toString().trim().toUpperCase() : "";
    var inputMp = d.kod_mp ? d.kod_mp.toString().trim().toUpperCase() : "";
    var inputKp = d.no_kp ? d.no_kp.toString().trim() : "";
    
    // PEMERIKSAAN ANTI-PENDUA (Cegah Lantikan Berganda)
    for (var i = 1; i < data.length; i++) {
      if (r > 0 && r === (i + 1)) continue; // Abaikan baris yang sedang dikemaskini
      
      var dbTahun = data[i][1] ? data[i][1].toString().trim() : "";
      var dbTahap = data[i][2] ? data[i][2].toString().trim().toUpperCase() : "";
      var dbPeringkat = data[i][3] ? data[i][3].toString().trim().toUpperCase() : "";
      var dbDaerah = data[i][4] ? data[i][4].toString().trim().toUpperCase() : "";
      var dbMp = data[i][5] ? data[i][5].toString().trim().toUpperCase() : "";
      var dbNoKp = data[i][7] ? data[i][7].toString().trim() : "";
      
      var namaJawatanPenuh = (dbPeringkat === "NEGERI") ? "KETUA PENTAKSIR NEGERI (KPN)" : "KETUA PENTAKSIR KAWASAN (KPKW)";
      var namaSubjekPenuh = mapSubjek[dbMp] ? mapSubjek[dbMp] : dbMp; 

      if (dbTahun === inputTahun && dbTahap === inputTahap && dbPeringkat === inputPeringkat && dbMp === inputMp && dbDaerah === inputDaerah) {
         lock.releaseLock();
         return { status: "FAILED", message: "Lantikan bagi jawatan " + namaJawatanPenuh + " untuk mata pelajaran " + namaSubjekPenuh + " (" + dbTahap + ") di daerah/negeri ini telah wujud!" };
      }

      if (dbTahun === inputTahun && dbNoKp !== "" && dbNoKp === inputKp) {
         lock.releaseLock();
         return { 
           status: "FAILED", 
           message: "Gagal! Guru ini (No. KP: " + inputKp + ") telah pun dilantik sebagai " + namaJawatanPenuh + " bagi mata pelajaran " + namaSubjekPenuh + " (" + dbTahap + ") pada sesi " + dbTahun + ". Seorang guru hanya dibenarkan memegang SATU jawatan pentaksir sahaja." 
         };
      }
    }

    var prefixSurat = d.peringkat === "NEGERI" ? "KPN" : "KPKw";
    var idMatrikTahap = (d.matrik_tahap === "TAHUN 4") ? "MP4" : "MP3";
    var idUnikSurat = inputTahun + prefixSurat + inputMp + idMatrikTahap;

    if (r > 0) {
      // PROSES KEMASKINI 
      sheet.getRange(r, 3).setValue(inputTahap);
      sheet.getRange(r, 4).setValue(inputPeringkat);
      sheet.getRange(r, 5).setValue(inputDaerah);
      sheet.getRange(r, 6).setValue(inputMp);
      sheet.getRange(r, 7).setValue(d.nama_guru ? d.nama_guru.toString().toUpperCase() : "");
      sheet.getRange(r, 8).setValue("'" + inputKp); 
      sheet.getRange(r, 9).setValue(d.kod_sekolah ? d.kod_sekolah.toString().toUpperCase() : "");
      sheet.getRange(r, 10).setValue(d.no_tel ? "'" + d.no_tel.toString() : ""); 
      sheet.getRange(r, 11).setValue(d.emel ? d.emel.toString() : "");
      sheet.getRange(r, 12).setValue(idUnikSurat);
    } else {
      // PROSES TAMBAH BARU (ID DIKIRA SECARA LIVE)
      var maxId = 0;
      for (var k = 1; k < data.length; k++) {
        var idVal = parseInt(data[k][0]) || 0;
        if (idVal > maxId) maxId = idVal;
      }
      var nextId = maxId + 1; // Sah, terkini, tidak akan duplicate
      
      sheet.appendRow([
        nextId, 
        inputTahun, 
        inputTahap, 
        inputPeringkat, 
        inputDaerah, 
        inputMp, 
        d.nama_guru ? d.nama_guru.toString().toUpperCase() : "", 
        "'" + inputKp, 
        d.kod_sekolah ? d.kod_sekolah.toString().toUpperCase() : "", 
        d.no_tel ? "'" + d.no_tel.toString() : "", 
        d.emel ? d.emel.toString() : "", 
        idUnikSurat
      ]);
    }
    SpreadsheetApp.flush();
    lock.releaseLock(); // SELESAI, BUKA KUNCI
    return { status: "SUCCESS" };
  } catch (e) {
    if (lock) lock.releaseLock();
    return { status: "FAILED", message: e.message };
  }
}

// =========================================================================
// 15. JANA SURAT LANTIKAN PENYELIA KAWASAN (PK)
// =========================================================================
function janaSuratLantikanPDF(noKp, tahun, tahap, isRawHtml) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var wsPk = ss.getSheetByName("tbpk");
    if (!wsPk) throw new Error("Tab sheet 'tbpk' tidak dijumpai.");
    var dataPk = wsPk.getDataRange().getValues();
    var headersPk = dataPk[0];
    var pkRow = null;
    
    for (var i = 1; i < dataPk.length; i++) {
      if (dataPk[i][headersPk.indexOf("no_kp")] == noKp && 
          dataPk[i][headersPk.indexOf("tahun")] == tahun && 
          dataPk[i][headersPk.indexOf("matrik_tahap")] == tahap) {
        pkRow = dataPk[i];
        break;
      }
    }
    
    if (!pkRow) throw new Error("Rekod Penyelia Kawasan tidak dijumpai di dalam pangkalan data.");
    
    var namaPk = pkRow[headersPk.indexOf("nama_pk")];
    var kodSekAsal = pkRow[headersPk.indexOf("kod_sek_asal")];
    
    var wsSekolah = ss.getSheetByName("tbsekolah");
    var dataSekolah = wsSekolah ? wsSekolah.getDataRange().getValues() : [];
    var namaSekAsal = kodSekAsal;
    var poskodAsal = "";
    var bandarAsal = "";
    
    if (dataSekolah.length > 0) {
      for (var i = 1; i < dataSekolah.length; i++) {
        var kdSek = dataSekolah[i][5] || dataSekolah[i][0];
        if (kdSek && kdSek.toString().trim().toUpperCase() === kodSekAsal.toString().trim().toUpperCase()) {
          namaSekAsal = dataSekolah[i][6] || dataSekolah[i][1] || kodSekAsal;
          poskodAsal = dataSekolah[i][8] || "";
          bandarAsal = dataSekolah[i][9] || dataSekolah[i][10] || "";
          break;
        }
      }
    }
    
    var wsMatrik = ss.getSheetByName("tbmatrik_operasi") || ss.getSheetByName("tbmatrik");
    var dataMatrik = wsMatrik ? wsMatrik.getDataRange().getValues() : [];
    
    var senaraiSekolahJagaan = [];
    var idBilikKebal = "";
    var HAD_MAKSIMUM_SEKOLAH = 9; // UBAH KUOTA KEPADA 9
    
    if (dataMatrik.length > 0) {
      var hdMatrik = dataMatrik[0].map(function(h) { return h.toString().trim().toUpperCase(); });
      var idxTh = hdMatrik.indexOf("TAHUN") > -1 ? hdMatrik.indexOf("TAHUN") : 1;
      var idxThp = hdMatrik.indexOf("MATRIK_TAHAP") > -1 ? hdMatrik.indexOf("MATRIK_TAHAP") : 4;
      var idxSek = hdMatrik.indexOf("KOD_SEKOLAH") > -1 ? hdMatrik.indexOf("KOD_SEKOLAH") : 2;
      var idxPk1 = hdMatrik.indexOf("NO_KP_PK1") > -1 ? hdMatrik.indexOf("NO_KP_PK1") : 6;
      var idxPk2 = hdMatrik.indexOf("NO_KP_PK2") > -1 ? hdMatrik.indexOf("NO_KP_PK2") : 7;
      var idxBk = hdMatrik.indexOf("ID_BILIK_KEBAL") > -1 ? hdMatrik.indexOf("ID_BILIK_KEBAL") : 5;

      for (var i = 1; i < dataMatrik.length; i++) {
        var mTahun = dataMatrik[i][idxTh];
        var mTahap = dataMatrik[i][idxThp];
        var mPk1 = dataMatrik[i][idxPk1];
        var mPk2 = dataMatrik[i][idxPk2];
        
        if (mTahun == tahun && mTahap.toString().toUpperCase() == tahap.toUpperCase() && (mPk1 == noKp || mPk2 == noKp)) {
          var mKodSek = dataMatrik[i][idxSek];
          
          var namaSekJagaan = mKodSek;
          for (var j = 1; j < dataSekolah.length; j++) {
            var kS = dataSekolah[j][5] || dataSekolah[j][0];
            if (kS && kS.toString().trim().toUpperCase() === mKodSek.toString().trim().toUpperCase()) {
              namaSekJagaan = dataSekolah[j][6] || dataSekolah[j][1];
              break;
            }
          }
          
          if (senaraiSekolahJagaan.length < HAD_MAKSIMUM_SEKOLAH) {
            senaraiSekolahJagaan.push(mKodSek + " - " + namaSekJagaan);
          }
          
          if (idBilikKebal === "") {
            idBilikKebal = dataMatrik[i][idxBk];
          }
        }
      }
    }
    
    // UBAH 1: BINA JADUAL HTML DENGAN NOMBOR (BIL)
    var jadualSekolahJagaan = "";
    if (senaraiSekolahJagaan.length > 0) {
      for(var r = 0; r < senaraiSekolahJagaan.length; r++) {
        jadualSekolahJagaan += "<tr><td style='text-align: center; font-weight: bold;'>" + (r + 1) + "</td><td><b>" + senaraiSekolahJagaan[r] + "</b></td></tr>";
      }
    } else {
      jadualSekolahJagaan = "<tr><td colspan='2' style='text-align: center; padding: 15px;'>TIADA REKOD SEKOLAH JAGAAN DITETAPKAN</td></tr>";
    }

    // UBAH 2: FORMAT BILIK KEBAL (NAMA DAERAH - NAMA BILIK KEBAL)
    var alamatBilikKebalPenuh = idBilikKebal || "BELUM DITETAPKAN";
    if (idBilikKebal) {
      var wsBilikKebal = ss.getSheetByName("tbbilikkebal");
      if (wsBilikKebal) {
        var dataBk = wsBilikKebal.getDataRange().getValues();
        for (var i = 1; i < dataBk.length; i++) {
          if (dataBk[i][0] && dataBk[i][0].toString().trim().toUpperCase() === idBilikKebal.toString().trim().toUpperCase()) {
            var namaBk = dataBk[i][1] ? dataBk[i][1].toString().trim() : "";
            var daerahBk = dataBk[i][2] ? dataBk[i][2].toString().trim().toUpperCase() : "";
            
            // Dapatkan nama daerah penuh
            var namaDaerahPenuh = getNamaDaerahFullBackend(daerahBk, ss); 
            alamatBilikKebalPenuh = (namaDaerahPenuh && namaDaerahPenuh !== "-" ? namaDaerahPenuh + " - " : "") + namaBk;
            break;
          }
        }
      }
    }

    var sheetSurat = ss.getSheetByName("tbsurat");
    var dataSurat = sheetSurat ? sheetSurat.getDataRange().getValues() : [];
    var konfigSurat = null;
    
    var idxIdUnik = headersPk.indexOf("no_rujukan_surat") > -1 ? headersPk.indexOf("no_rujukan_surat") : 12;
    var idUnikSurat = pkRow[idxIdUnik] ? pkRow[idxIdUnik].toString().trim().toUpperCase() : "";

    for (var s = 1; s < dataSurat.length; s++) {
      var idRow = dataSurat[s][0] ? dataSurat[s][0].toString().trim().toUpperCase() : "";
      if (idUnikSurat !== "" && idRow === idUnikSurat) {
        konfigSurat = {
          nama_surat: dataSurat[s][1],
          rujukan_fail: dataSurat[s][2],
          tarikh_surat: dataSurat[s][3], 
          nama_pengarah: dataSurat[s][4],
          pingat_pengarah: dataSurat[s][5],
          jawatan_pengarah: dataSurat[s][6],
          nama_jabatan: dataSurat[s][7],
          id_gambar_sain: dataSurat[s][8],
          tarikh_ujian: dataSurat[s][9]
        };
        break;
      }
    }
    
    if (!konfigSurat) {
      for (var s = 1; s < dataSurat.length; s++) {
        var idR = dataSurat[s][0] ? dataSurat[s][0].toString().trim().toUpperCase() : "";
        if (idR.indexOf("PK") > -1 && (idR.indexOf(tahap.replace(/\s+/g,'')) > -1)) {
          konfigSurat = {
            nama_surat: dataSurat[s][1],
            rujukan_fail: dataSurat[s][2],
            tarikh_surat: dataSurat[s][3], 
            nama_pengarah: dataSurat[s][4],
            pingat_pengarah: dataSurat[s][5],
            jawatan_pengarah: dataSurat[s][6],
            nama_jabatan: dataSurat[s][7],
            id_gambar_sain: dataSurat[s][8],
            tarikh_ujian: dataSurat[s][9]
          };
          break;
        }
      }
    }

    if (!konfigSurat) throw new Error("Tetapan konfigurasi Surat PK tidak dijumpai dalam tbsurat.");

    var htmlSainPengarah = (konfigSurat.id_gambar_sain && konfigSurat.id_gambar_sain.toString().trim() !== "") 
      ? '<img src="https://drive.google.com/uc?export=view&id=' + konfigSurat.id_gambar_sain.toString().trim() + '" style="height: 90px; width: auto; margin-top: -15px; margin-bottom: -30px;">' 
      : "t.t";

    var htmlTemplate = HtmlService.createHtmlOutputFromFile('surat_lantikan_pk').getContent();
    
    var bM = ["Januari","Februari","Mac","April","Mei","Jun","Julai","Ogos","September","Oktober","November","Disember"];
    var tK = new Date(); 
    var tarikhAuto = tK.getDate() + " " + bM[tK.getMonth()] + " " + tK.getFullYear();
    var tarikhPaparan = (konfigSurat.tarikh_surat && konfigSurat.tarikh_surat.toString().trim() !== "") ? konfigSurat.tarikh_surat.toString().trim() : tarikhAuto;

    var pingatTek = konfigSurat.pingat_pengarah ? (", " + konfigSurat.pingat_pengarah) : "";

    var htmlDisesuaikan = htmlTemplate
      .replace(/{{RUJUKAN}}/g, konfigSurat.rujukan_fail || "-") 
      .replace(/{{NAMA_SURAT}}/g, konfigSurat.nama_surat || "MATRIKS PEMBELAJARAN") 
      .replace(/{{TARIKH}}/g, tarikhPaparan)
      .replace(/{{NAMA}}/g, (namaPk || "").toUpperCase())
      .replace(/{{NO_KP}}/g, noKp)
      .replace(/{{SEKOLAH_ASAL}}/g, (namaSekAsal || "").toUpperCase())
      .replace(/{{POSKOD}}/g, poskodAsal)
      .replace(/{{BANDAR}}/g, (bandarAsal || "").toUpperCase())
      .replace(/{{TAHUN}}/g, tahun)
      .replace(/{{JAWATAN_BESAR}}/g, "PENYELIA KAWASAN") 
      .replace(/{{JAWATAN_TITLE}}/g, "Penyelia Kawasan")
      .replace(/{{TARIKH_UJIAN}}/g, konfigSurat.tarikh_ujian || "AKAN DIMAKLUMKAN")
      .replace(/{{TR_SEKOLAH_JAGAAN}}/g, jadualSekolahJagaan)
      .replace(/{{BILIK_KEBAL}}/g, alamatBilikKebalPenuh.toUpperCase())
      .replace(/{{SAIN_PENGARAH}}/g, htmlSainPengarah)
      .replace(/{{NAMA_PENGARAH}}/g, (konfigSurat.nama_pengarah || "").toUpperCase())
      .replace(/{{PINGAT_PENGARAH}}/g, pingatTek)
      .replace(/{{JAWATAN_PENGARAH}}/g, konfigSurat.jawatan_pengarah || "")
      .replace(/{{NAMA_JABATAN}}/g, konfigSurat.nama_jabatan || "");
      
    if (isRawHtml) return htmlDisesuaikan;

    var blob = Utilities.newBlob(htmlDisesuaikan, MimeType.HTML, "temp.html").getAs(MimeType.PDF);
    var cleanName = (namaPk).replace(/[^a-zA-Z0-9 ]/g, "").replace(/\s+/g, "_");
    var namaFail = "Lantikan_PK_" + cleanName + "_" + tahun + ".pdf";
    blob.setName(namaFail);
    
    return {
      status: "SUCCESS",
      base64: Utilities.base64Encode(blob.getBytes()),
      filename: namaFail
    };
    
  } catch (error) {
    if (isRawHtml) return "Ralat: " + error.toString();
    return {
      status: "ERROR",
      message: "Ralat menjana PDF Penyelia Kawasan: " + error.toString()
    };
  }
}

// =========================================================================
// 16. PENJANA SURAT KETUA PENGAWAS (KP)
// =========================================================================
function janaSuratLantikanKPPDF(noKp, tahun, tahap, isRawHtml) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var sheetKp = ss.getSheetByName('tbkp');
    var dataKp = sheetKp ? sheetKp.getDataRange().getValues() : [];
    var rekodKp = null;
    var hdKp = dataKp[0].map(function(h) { return h.toString().trim().toUpperCase(); });
    
    var iTh = hdKp.indexOf("TAHUN") > -1 ? hdKp.indexOf("TAHUN") : 1;
    var iThp = hdKp.indexOf("MATRIK_TAHAP") > -1 ? hdKp.indexOf("MATRIK_TAHAP") : 4;
    var iKp = hdKp.indexOf("NO_KP") > -1 ? hdKp.indexOf("NO_KP") : 7;
    var iAsal = hdKp.indexOf("KOD_SEK_ASAL") > -1 ? hdKp.indexOf("KOD_SEK_ASAL") : 2;
    var iTtp = hdKp.indexOf("KOD_SEK_TETAP") > -1 ? hdKp.indexOf("KOD_SEK_TETAP") : 12;
    var iNm = hdKp.indexOf("NAMA_GURU") > -1 ? hdKp.indexOf("NAMA_GURU") : 8;
    var iSt = hdKp.indexOf("STATUS_PENETAPAN") > -1 ? hdKp.indexOf("STATUS_PENETAPAN") : 13;
    var iUnik = hdKp.indexOf("NO_RUJUKAN_SURAT") > -1 ? hdKp.indexOf("NO_RUJUKAN_SURAT") : 14; 

    for (var i = 1; i < dataKp.length; i++) {
      if (dataKp[i][iTh].toString().trim() == tahun && dataKp[i][iThp].toString().trim().toUpperCase() == tahap.toUpperCase() && dataKp[i][iKp].toString().trim() == noKp) {
        rekodKp = { 
          tahun: dataKp[i][iTh], 
          kod_sek_asal: dataKp[i][iAsal], 
          kod_sek_tetap: dataKp[i][iTtp], 
          nama: dataKp[i][iNm], 
          no_kp: dataKp[i][iKp], 
          status: dataKp[i][iSt], 
          id_unik_surat: dataKp[i][iUnik] 
        }; 
        break;
      }
    }
    
    if (!rekodKp) return isRawHtml ? "Ralat: Rekod KP tidak dijumpai." : { status: "FAILED", message: "Rekod KP tidak dijumpai." };
    if (rekodKp.status.toString().toUpperCase() !== "DISAHKAN") return isRawHtml ? "Ralat: Status masih CADANGAN." : { status: "FAILED", message: "Status masih CADANGAN." };

    var sheetSurat = ss.getSheetByName('tbsurat');
    var dataSurat = sheetSurat ? sheetSurat.getDataRange().getValues() : [];
    var konfigSurat = null;
    var idDicari = (rekodKp.id_unik_surat || "").toString().trim().toUpperCase();

    for (var j = 1; j < dataSurat.length; j++) {
      var idRow = dataSurat[j][0] ? dataSurat[j][0].toString().trim().toUpperCase() : "";
      if (idDicari !== "" && idRow === idDicari) {
        konfigSurat = { 
          nama_surat: dataSurat[j][1], 
          rujukan_fail: dataSurat[j][2], 
          tarikh_surat: dataSurat[j][3], 
          tarikh_ujian: dataSurat[j][9], 
          nama_pengarah: dataSurat[j][4], 
          pingat_pengarah: dataSurat[j][5], 
          jawatan_pengarah: dataSurat[j][6], 
          nama_jabatan: dataSurat[j][7],
          id_gambar_sain: dataSurat[j][8] 
        }; 
        break;
      }
    }
    
    if (!konfigSurat) {
      for (var j = 1; j < dataSurat.length; j++) {
         var idR = dataSurat[j][0] ? dataSurat[j][0].toString().trim().toUpperCase() : "";
         if (idR.indexOf("KP") > -1 && (idR.indexOf(tahap.replace(/\s+/g,'')) > -1)) {
           konfigSurat = { 
             nama_surat: dataSurat[j][1], 
             rujukan_fail: dataSurat[j][2], 
             tarikh_surat: dataSurat[j][3],
             tarikh_ujian: dataSurat[j][9], 
             nama_pengarah: dataSurat[j][4], 
             pingat_pengarah: dataSurat[j][5], 
             jawatan_pengarah: dataSurat[j][6], 
             nama_jabatan: dataSurat[j][7],
             id_gambar_sain: dataSurat[j][8] 
           }; 
           break;
         }
      }
    }
    if (!konfigSurat) return isRawHtml ? "Ralat: Tetapan konfigurasi Surat KP tidak dijumpai." : { status: "FAILED", message: "Tetapan konfigurasi Surat KP tidak dijumpai dalam tbsurat." };

    var sheetSek = ss.getSheetByName('tbsekolah');
    var dataSek = sheetSek ? sheetSek.getDataRange().getValues() : [];
    var namaSekAsal = rekodKp.kod_sek_asal, poskodSekAsal = "", bandarSekAsal = "", namaSekPusat = rekodKp.kod_sek_tetap;
    
    for (var s = 1; s < dataSek.length; s++) {
      if (dataSek[s][5]) {
        var kd = dataSek[s][5].toString().trim().toUpperCase();
        if (kd === rekodKp.kod_sek_asal.toUpperCase()) { 
          namaSekAsal = dataSek[s][6] || rekodKp.kod_sek_asal; 
          poskodSekAsal = dataSek[s][8] || ""; 
          bandarSekAsal = dataSek[s][9] || ""; 
        }
        if (kd === rekodKp.kod_sek_tetap.toUpperCase()) { 
          namaSekPusat = kd + " " + dataSek[s][6]; 
        }
      }
    }

    var sheetMatrik = ss.getSheetByName('tbmatrik_operasi');
    var dataMatrik = sheetMatrik ? sheetMatrik.getDataRange().getValues() : [];
    var noKpPk1 = "", idBilikKebal = "";
    
    for (var m = 1; m < dataMatrik.length; m++) {
      if (dataMatrik[m][1].toString().trim() == tahun && dataMatrik[m][4].toString().trim().toUpperCase() == tahap.toUpperCase() && dataMatrik[m][2].toString().trim().toUpperCase() == rekodKp.kod_sek_tetap.toUpperCase()) {
        noKpPk1 = dataMatrik[m][6] ? dataMatrik[m][6].toString().trim() : "";
        idBilikKebal = dataMatrik[m][5] ? dataMatrik[m][5].toString().trim() : ""; 
        break;
      }
    }

    var paparanBilikKebal = idBilikKebal || "BELUM DITETAPKAN";
    var sheetBk = ss.getSheetByName('tbbilikkebal');
    var dataBk = sheetBk ? sheetBk.getDataRange().getValues() : [];
    for (var b = 1; b < dataBk.length; b++) {
      if (dataBk[b][0] && dataBk[b][0].toString().trim().toUpperCase() === idBilikKebal.toUpperCase()) { 
        paparanBilikKebal = dataBk[b][1] || idBilikKebal; 
        break; 
      }
    }

    var namaPk1 = "BELUM DITETAPKAN", kodSekPk1 = "", namaSekPk1 = "BELUM DITETAPKAN";
    if (noKpPk1 !== "") {
      var sheetPk = ss.getSheetByName('tbpk');
      var dataPkT = sheetPk ? sheetPk.getDataRange().getValues() : [];
      for (var p = 1; p < dataPkT.length; p++) {
        if (dataPkT[p][1].toString().trim() == tahun && dataPkT[p][4].toString().trim().toUpperCase() == tahap.toUpperCase() && dataPkT[p][8].toString().trim() == noKpPk1) {
          namaPk1 = dataPkT[p][7] || noKpPk1; 
          kodSekPk1 = dataPkT[p][2] || ""; 
          break;
        }
      }
      if (kodSekPk1 !== "") {
        namaSekPk1 = kodSekPk1;
        for (var sk = 1; sk < dataSek.length; sk++) {
          if (dataSek[sk][5] && dataSek[sk][5].toString().trim().toUpperCase() === kodSekPk1.toUpperCase()) { 
            namaSekPk1 = dataSek[sk][6]; 
            break; 
          }
        }
      }
    }

    var htmlTemplate = HtmlService.createHtmlOutputFromFile('surat_lantikan_kp').getContent();
    
    var bM = ["Januari","Februari","Mac","April","Mei","Jun","Julai","Ogos","September","Oktober","November","Disember"];
    var tK = new Date(); 
    var tarikhAuto = tK.getDate() + " " + bM[tK.getMonth()] + " " + tK.getFullYear();
    var tarikhPaparan = (konfigSurat.tarikh_surat && konfigSurat.tarikh_surat.toString().trim() !== "") ? konfigSurat.tarikh_surat.toString().trim() : tarikhAuto;

    var htmlSainPengarahKP = (konfigSurat.id_gambar_sain && konfigSurat.id_gambar_sain.trim() !== "") 
      ? '<img src="https://drive.google.com/uc?export=view&id=' + konfigSurat.id_gambar_sain.trim() + '" style="height: 90px; width: auto; margin-top: -15px; margin-bottom: -30px;">' 
      : "t.t";

    htmlTemplate = htmlTemplate.replace(/{{RUJUKAN}}/g, konfigSurat.rujukan_fail || "-");
    htmlTemplate = htmlTemplate.replace(/{{NAMA_SURAT}}/g, konfigSurat.nama_surat || "MATRIKS PEMBELAJARAN"); 
    htmlTemplate = htmlTemplate.replace(/{{TARIKH}}/g, tarikhPaparan);
    htmlTemplate = htmlTemplate.replace(/{{NAMA}}/g, rekodKp.nama || "");
    htmlTemplate = htmlTemplate.replace(/{{NO_KP}}/g, rekodKp.no_kp || "");
    htmlTemplate = htmlTemplate.replace(/{{SEKOLAH_ASAL}}/g, namaSekAsal || "");
    htmlTemplate = htmlTemplate.replace(/{{POSKOD}}/g, poskodSekAsal || "");
    htmlTemplate = htmlTemplate.replace(/{{BANDAR}}/g, bandarSekAsal || "");
    htmlTemplate = htmlTemplate.replace(/{{TAHUN}}/g, rekodKp.tahun || "");
    
    htmlTemplate = htmlTemplate.replace(/{{JAWATAN_BESAR}}/g, "KETUA PENGAWAS");
    htmlTemplate = htmlTemplate.replace(/{{JAWATAN_TITLE}}/g, "Ketua Pengawas");
    htmlTemplate = htmlTemplate.replace(/{{TARIKH_UJIAN}}/g, konfigSurat.tarikh_ujian || "AKAN DIMAKLUMKAN");
    
    htmlTemplate = htmlTemplate.replace(/{{PUSAT_PENUGASAN}}/g, namaSekPusat || "");
    htmlTemplate = htmlTemplate.replace(/{{BILIK_KEBAL}}/g, paparanBilikKebal);
    
    htmlTemplate = htmlTemplate.replace(/{{PENYELIA_KAWASAN}}/g, namaPk1);
    htmlTemplate = htmlTemplate.replace(/{{SEKOLAH_ASAL_PENYELIA_KAWASAN}}/g, namaSekPk1);

    htmlTemplate = htmlTemplate.replace(/{{SAIN_PENGARAH}}/g, htmlSainPengarahKP); 
    htmlTemplate = htmlTemplate.replace(/{{NAMA_PENGARAH}}/g, konfigSurat.nama_pengarah || "");
    htmlTemplate = htmlTemplate.replace(/{{PINGAT_PENGARAH}}/g, konfigSurat.pingat_pengarah ? ", " + konfigSurat.pingat_pengarah : "");
    htmlTemplate = htmlTemplate.replace(/{{JAWATAN_PENGARAH}}/g, konfigSurat.jawatan_pengarah || "");
    htmlTemplate = htmlTemplate.replace(/{{NAMA_JABATAN}}/g, konfigSurat.nama_jabatan || "");
    
    if (isRawHtml) return htmlTemplate;

    var blobPDF = Utilities.newBlob(htmlTemplate, MimeType.HTML).getAs(MimeType.PDF);
    var cleanName = (rekodKp.nama).replace(/[^a-zA-Z0-9 ]/g, "").replace(/\s+/g, "_");
    blobPDF.setName("Surat_Lantikan_KP_" + cleanName + "_" + tahun + ".pdf");

    return { 
      status: "SUCCESS", 
      base64: Utilities.base64Encode(blobPDF.getBytes()), 
      filename: blobPDF.getName() 
    };

  } catch (e) { 
    if (isRawHtml) return "Ralat: " + e.message;
    return { status: "FAILED", message: "Ralat Penjanaan PDF HTML: " + e.message }; 
  }
}

// =========================================================================
// 17. PENJANA SURAT IKUT PUSAT PEPERIKSAAN
// =========================================================================
function janaSuratPusatPDF(kodSekolah, tahun, tahap) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var sheetMatrik = ss.getSheetByName("tbmatrik_operasi");
    if (!sheetMatrik) return { status: "ERROR", message: "Tab tbmatrik_operasi tidak dijumpai." };
    var dataMatrik = sheetMatrik.getDataRange().getValues();
    
    var noKpPk1 = "", noKpPk2 = "";
    for (var i = 1; i < dataMatrik.length; i++) {
      if (dataMatrik[i][2] && dataMatrik[i][2].toString().trim().toUpperCase() === kodSekolah.toUpperCase() &&
          dataMatrik[i][1].toString().trim() == tahun && 
          dataMatrik[i][4].toString().trim().toUpperCase() === tahap.toUpperCase()) {
        noKpPk1 = dataMatrik[i][6] ? dataMatrik[i][6].toString().trim() : "";
        noKpPk2 = dataMatrik[i][7] ? dataMatrik[i][7].toString().trim() : "";
        break;
      }
    }
    
    var sheetKp = ss.getSheetByName("tbkp");
    var dataKp = sheetKp ? sheetKp.getDataRange().getValues() : [];
    var noKpKp = "";
    for (var k = 1; k < dataKp.length; k++) {
      if (dataKp[k][12] && dataKp[k][12].toString().trim().toUpperCase() === kodSekolah.toUpperCase() &&
          dataKp[k][1].toString().trim() == tahun &&
          dataKp[k][4].toString().trim().toUpperCase() === tahap.toUpperCase() &&
          dataKp[k][13].toString().trim().toUpperCase() === 'DISAHKAN') {
        noKpKp = dataKp[k][7] ? dataKp[k][7].toString().trim() : "";
        break;
      }
    }

    var htmlKeseluruhan = "";
    var senaraiPetugas = [];
    
    if (noKpPk1 !== "") senaraiPetugas.push({ jenis: 'PK', no_kp: noKpPk1 });
    if (noKpPk2 !== "") senaraiPetugas.push({ jenis: 'PK', no_kp: noKpPk2 });
    if (noKpKp !== "") senaraiPetugas.push({ jenis: 'KP', no_kp: noKpKp });
    
    if (senaraiPetugas.length === 0) {
      return { status: "ERROR", message: "Tiada petugas (PK/KP) yang sah untuk pusat ini." };
    }

    for (var j = 0; j < senaraiPetugas.length; j++) {
      var petugas = senaraiPetugas[j];
      var htmlSemasa = "";
      
      if (petugas.jenis === 'PK') {
        htmlSemasa = janaSuratLantikanPDF(petugas.no_kp, tahun, tahap, true);
      } else {
        htmlSemasa = janaSuratLantikanKPPDF(petugas.no_kp, tahun, tahap, true);
      }
      
      if (htmlSemasa && typeof htmlSemasa === 'string' && htmlSemasa.indexOf("Ralat:") === -1) {
        htmlKeseluruhan += htmlSemasa;
        if (j < senaraiPetugas.length - 1) {
          htmlKeseluruhan += "<div style='page-break-after: always;'></div>";
        }
      } else {
        return { status: "ERROR", message: "Gagal menjana surat untuk No KP: " + petugas.no_kp + " (" + htmlSemasa + ")" };
      }
    }
    
    var blob = Utilities.newBlob(htmlKeseluruhan, MimeType.HTML).setName("Surat_Petugas_" + kodSekolah + ".pdf");
    var pdfBlob = blob.getAs(MimeType.PDF);
    var base64Data = Utilities.base64Encode(pdfBlob.getBytes());
    
    return { 
      status: "SUCCESS", 
      base64: base64Data, 
      filename: "Surat_Petugas_" + kodSekolah + ".pdf" 
    };

  } catch (error) {
    return { status: "ERROR", message: error.toString() };
  }
}

// ==========================================================
// 18. FUNGSI HELPER: AMBIL SENARAI DAERAH UNTUK DROPDOWN LAPORAN
// ==========================================================
function dapatkanSenaraiDaerahUntukLaporan() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ws = ss.getSheetByName("tbdaerah");
    if (!ws) return [];
    
    var data = ws.getDataRange().getValues();
    if (data.length > 0) data.shift();
    
    var senaraiPPD = [];
    var semakPendua = {};
    
    data.forEach(function(row) {
      var kodPpd = row[5] ? row[5].toString().trim().toUpperCase() : "";
      var namaPpd = row[2] ? row[2].toString().trim().toUpperCase() : "";
      
      if (kodPpd !== "" && !semakPendua[kodPpd]) {
        senaraiPPD.push({
          kod: kodPpd,
          nama: namaPpd
        });
        semakPendua[kodPpd] = true;
      }
    });
    
    senaraiPPD.sort(function(a, b) {
      return a.kod.localeCompare(b.kod);
    });
    
    return senaraiPPD;
  } catch(e) {
    return [];
  }
}

// ==========================================================
// FUNGSI HELPER: SEMAK JANTINA DARIPADA NO KAD PENGENALAN
// ==========================================================
function tentukanJantinaDariKP(noKp) {
  if (!noKp) return "-";
  var nomborSahaja = noKp.toString().replace(/[^0-9]/g, '');
  if (!nomborSahaja || nomborSahaja.length === 0) return "-";
  
  var digitAkhir = parseInt(nomborSahaja.slice(-1), 10);
  return (digitAkhir % 2 === 0) ? 'P' : 'L';
}

// ==========================================================
// FUNGSI UTAMA: JANA BORANG SPP2 PDF MENGGUNAKAN TEMPLATE SPP2
// ==========================================================
function janaBorangSPP2PDF(kodDaerah, jenisPeperiksaan, tahun) {
  try {
    jenisPeperiksaan = jenisPeperiksaan || "MATRIKS PEMBELAJARAN TAHUN 4";
    tahun = tahun || new Date().getFullYear().toString(); 

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var wsSekolahMaster = ss.getSheetByName("tbsekolah");
    var mapSekolah = {};
    var mapCalon = {};
    if (wsSekolahMaster) {
       var dataSekMaster = wsSekolahMaster.getDataRange().getValues();
       for(var s = 1; s < dataSekMaster.length; s++){
          var kodS = dataSekMaster[s][5] ? dataSekMaster[s][5].toString().trim().toUpperCase() : "";
          var namaS = dataSekMaster[s][6] ? dataSekMaster[s][6].toString().trim().toUpperCase() : "";
          var jumC = dataSekMaster[s][23] ? dataSekMaster[s][23].toString().trim() : "0";
          if(kodS) {
            mapSekolah[kodS] = namaS;
            mapCalon[kodS] = jumC;
          }
       }
    }

    var wsDaerah = ss.getSheetByName("tbdaerah");
    var mapDaerah = {};
    if (wsDaerah) {
      var dataDaerah = wsDaerah.getDataRange().getValues();
      for(var d = 1; d < dataDaerah.length; d++){
         var kodD = dataDaerah[d][5] ? dataDaerah[d][5].toString().trim().toUpperCase() : "";
         var namaD = dataDaerah[d][2] ? dataDaerah[d][2].toString().trim().toUpperCase() : "";
         if(kodD) mapDaerah[kodD] = namaD;
      }
    }

    var wsBK = ss.getSheetByName("tbbilikkebal");
    var mapBK = {};
    if (wsBK) {
      var dataBK = wsBK.getDataRange().getValues();
      for(var b = 1; b < dataBK.length; b++){
         var idBK = dataBK[b][0] ? dataBK[b][0].toString().trim().toUpperCase() : ""; 
         var namaBK = dataBK[b][1] ? dataBK[b][1].toString().trim().toUpperCase() : "";
         if(idBK) mapBK[idBK] = namaBK;
      }
    }

    var wsMatrik = ss.getSheetByName("tbmatrik_operasi"); 
    var dataMatrik = wsMatrik ? wsMatrik.getDataRange().getValues() : [];
    if (dataMatrik.length > 0) dataMatrik.shift(); 

    var wsKP = ss.getSheetByName("tbkp");
    var dataKP = wsKP ? wsKP.getDataRange().getValues() : [];
    if (dataKP.length > 0) dataKP.shift();

    var wsPK = ss.getSheetByName("tbpk");
    var dataPK = wsPK ? wsPK.getDataRange().getValues() : [];
    if (dataPK.length > 0) dataPK.shift();

    // 1. PANGGIL DATA PENGAWAS (tbp)
    var wsPgws = ss.getSheetByName("tbp");
    var dataPgws = wsPgws ? wsPgws.getDataRange().getValues() : [];
    if (dataPgws.length > 0) dataPgws.shift();

    var pusatTapis = dataMatrik.filter(function(row) {
      var tahunRow = row[1] ? row[1].toString().trim() : "";
      var kodDaerahRow = row[3] ? row[3].toString().trim().toUpperCase() : "";
      var matrikTahapRow = row[4] ? row[4].toString().toUpperCase().trim() : "";

      var padanDaerah = (kodDaerahRow === kodDaerah.toUpperCase() || kodDaerah === "SEMUA");
      var padanTahun = (tahunRow === tahun.toString().trim());
      
      var padanTahap = true;
      if (jenisPeperiksaan.toUpperCase().indexOf("3") !== -1) {
        padanTahap = (matrikTahapRow.indexOf("3") !== -1);
      } else if (jenisPeperiksaan.toUpperCase().indexOf("4") !== -1) {
        padanTahap = (matrikTahapRow.indexOf("4") !== -1);
      }

      return padanDaerah && padanTahun && padanTahap;
    });

    if (pusatTapis.length === 0) {
      return { 
        success: false, 
        message: `Tiada pusat peperiksaan dijumpai bagi daerah (${kodDaerah}), tahap (${jenisPeperiksaan}) untuk tahun ${tahun}.` 
      };
    }

    var tarikhCetakan = Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "dd/MM/yyyy");
    var htmlCardsArray = [];

    pusatTapis.forEach(function(pusat) {
      var kodPusat = pusat[2] ? pusat[2].toString().trim().toUpperCase() : "";       
      var namaPusatFull = mapSekolah[kodPusat] ? (kodPusat + " - " + mapSekolah[kodPusat]) : kodPusat;
      
      var kodDaerahRow = pusat[3] ? pusat[3].toString().trim().toUpperCase() : "";
      var namaDaerahFull = mapDaerah[kodDaerahRow] ? (kodDaerahRow + " - " + mapDaerah[kodDaerahRow]) : kodDaerahRow;     
      
      var idBKRow = pusat[5] ? pusat[5].toString().trim().toUpperCase() : ""; 
      var infoBilikKebal = (idBKRow && idBKRow !== "-") ? (mapBK[idBKRow] ? (idBKRow + " - " + mapBK[idBKRow]) : idBKRow) : "-";
      
      var jumlahCalon = mapCalon[kodPusat] || "0";
      
      var icPk1 = pusat[6] ? pusat[6].toString().trim() : "";
      var icPk2 = pusat[7] ? pusat[7].toString().trim() : "";

      var senaraiPetugasPusat = [];

      // DATA PENYELIA KAWASAN
      [icPk1, icPk2].forEach(function(ic) {
        if(ic !== "") {
          var pkFound = dataPK.find(function(r) { return r[8].toString().trim() === ic && r[1].toString() === tahun.toString(); });
          if(pkFound) {
            var kodSekPk = pkFound[2] ? pkFound[2].toString().trim().toUpperCase() : "";
            senaraiPetugasPusat.push({
              jawatan: "PENYELIA KAWASAN",
              nama: pkFound[7] || "", 
              sekolahAsal: mapSekolah[kodSekPk] ? (kodSekPk + " - " + mapSekolah[kodSekPk]) : kodSekPk,
              noKp: pkFound[8] || "", 
              noTel: pkFound[9] || "" 
            });
          }
        }
      });

      // DATA KETUA PENGAWAS
      dataKP.forEach(function(kp) {
        var dbTahunKp = kp[1] ? kp[1].toString().trim() : "";
        var status = kp[13] ? kp[13].toString().trim().toUpperCase() : "";
        var kodSekTetap = kp[12] ? kp[12].toString().trim().toUpperCase() : "";
        var tahapKp = kp[4] ? kp[4].toString().trim().toUpperCase() : "";
        
        if (dbTahunKp === tahun.toString() && kodSekTetap === kodPusat && status === 'DISAHKAN') {
          if ((jenisPeperiksaan.indexOf("3") !== -1 && tahapKp.indexOf("3") !== -1) || 
              (jenisPeperiksaan.indexOf("4") !== -1 && tahapKp.indexOf("4") !== -1)) {
              
            var kodSekKp = kp[2] ? kp[2].toString().trim().toUpperCase() : "";
            senaraiPetugasPusat.push({
              jawatan: "KETUA PENGAWAS",
              nama: kp[6] || "",      
              sekolahAsal: mapSekolah[kodSekKp] ? (kodSekKp + " - " + mapSekolah[kodSekKp]) : kodSekKp,
              noKp: kp[7] || "",      
              noTel: kp[8] || ""      
            });
          }
        }
      });

      // 2. MASUKKAN DATA PENGAWAS
      dataPgws.forEach(function(pgws) {
        var dbTahunPgws = pgws[1] ? pgws[1].toString().trim() : "";
        var kodSekPgws = pgws[2] ? pgws[2].toString().trim().toUpperCase() : "";
        var tahapPgws = pgws[4] ? pgws[4].toString().trim().toUpperCase() : "";
        
        // Tapis ikut tahun dan kod pusat yang sedang diproses
        if (dbTahunPgws === tahun.toString() && kodSekPgws === kodPusat) {
          if ((jenisPeperiksaan.toUpperCase().indexOf("3") !== -1 && tahapPgws.indexOf("3") !== -1) || 
              (jenisPeperiksaan.toUpperCase().indexOf("4") !== -1 && tahapPgws.indexOf("4") !== -1)) {
              
            senaraiPetugasPusat.push({
              jawatan: "PENGAWAS",
              nama: pgws[5] || "",      
              sekolahAsal: mapSekolah[kodSekPgws] ? (kodSekPgws + " - " + mapSekolah[kodSekPgws]) : kodSekPgws,
              noKp: pgws[6] || "",      
              noTel: pgws[8] || "" // Index 8 ialah NO_TEL
            });
          }
        }
      });

      var rowsHTML = "";
      if (senaraiPetugasPusat.length > 0) {
        senaraiPetugasPusat.forEach(function(p, idx) {
          var jantina = tentukanJantinaDariKP(p.noKp);
          rowsHTML += `
            <tr>
              <td class="text-center">${idx + 1}</td>
              <td>${p.jawatan}</td>
              <td>${p.nama}</td>
              <td>${p.sekolahAsal}</td>
              <td class="text-center">${jantina}</td>
              <td class="text-center" style="white-space: nowrap;">${p.noTel}</td>
            </tr>
          `;
        });
      } else {
        rowsHTML = `<tr><td colspan="6" class="text-center" style="padding:15px; color:#666;">Tiada Petugas Didaftarkan Bagi Pusat Ini</td></tr>`;
      }

      var cardHTML = `
        <div class="report-card">
          <div>
            <div class="top-header-bar">
              <div>KOD PUSAT PEPERIKSAAN: <span>${kodPusat}</span></div>
              <div>BORANG SPP2</div>
            </div>

            <div class="header-center">
              <div class="divider">|</div>
              <div class="title-line">LEMBAGA PEPERIKSAAN</div>
              <div class="title-line">KEMENTERIAN PENDIDIKAN</div>
              <div class="sub-line">
                PEPERIKSAAN : ${jenisPeperiksaan.toUpperCase()} &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; TAHUN : <span>${tahun}</span>
              </div>
              <div class="sub-line">SENARAI PETUGAS YANG DILANTIK</div>
            </div>

            <div class="info-container">
              <table class="info-table">
                <tr>
                  <td style="width: 275px;">KOD DAN NAMA PUSAT PEPERIKSAAN</td>
                  <td style="width: 18px;">:</td>
                  <td><span>${namaPusatFull}</span></td>
                </tr>
              </table>

              <div style="margin-top: 15px;">
                <table class="info-table">
                  <tr>
                    <td><span class="underline-text">INFO CALON</span></td>
                    <td style="text-align: right;">DAERAH : <span>${namaDaerahFull}</span></td>
                  </tr>
                  <tr>
                    <td>JUMLAH CALON : <span>${jumlahCalon}</span> ORANG</td>
                    <td style="text-align: right;">BILIK KEBAL : <span>${infoBilikKebal}</span></td>
                  </tr>
                </table>
              </div>
            </div>

            <table class="data-table">
              <thead>
                <tr>
                  <th style="width: 4%;">BIL</th>
                  <th style="width: 20%;">JAWATAN PETUGAS</th>
                  <th style="width: 26%;">NAMA PETUGAS</th>
                  <th style="width: 26%;">NAMA SEKOLAH PETUGAS</th>
                  <th style="width: 9%;">JANTINA</th>
                  <th style="width: 15%; white-space: nowrap;">NO TELEFON</th>
                </tr>
              </thead>
              <tbody>
                ${rowsHTML}
              </tbody>
            </table>
          </div>

          <div class="footer-section">
            TARIKH <span class="underline-text">CETAKAN</span> : <span>${tarikhCetakan}</span>
          </div>
        </div>
      `;

      htmlCardsArray.push(cardHTML);
    });

    var htmlTemplate = HtmlService.createTemplateFromFile('TemplateSPP2');
    htmlTemplate.htmlCards = htmlCardsArray.join('');
    
    var htmlContent = htmlTemplate.evaluate().getContent();
    var blob = Utilities.newBlob(htmlContent, MimeType.HTML).getAs(MimeType.PDF);
    
    blob.setName(`Borang_SPP2_${kodDaerah}_${tahun}_${tarikhCetakan.replace(/\//g, '-')}.pdf`);

    return {
      success: true,
      pdfBase64: Utilities.base64Encode(blob.getBytes()),
      filename: blob.getName()
    };

  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// =========================================================================
// 19. PENJANA SURAT LANTIKAN KETUA PENTAKSIR NEGERI & KAWASAN (KPN & KPKw)
// =========================================================================
function janaSuratLantikanKpnKpkwPDF(noKp, tahun, tahap, isRawHtml) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Tarik Data KPN/KPKw
    var sheetKpn = ss.getSheetByName('tbkpn_kpkw');
    var dataKpn = sheetKpn ? sheetKpn.getDataRange().getValues() : [];
    var rekod = null;
    var hd = dataKpn[0].map(function(h) { return h.toString().trim().toLowerCase(); });
    
    var iTahun = hd.indexOf("tahun");
    var iTahap = hd.indexOf("matrik_tahap");
    var iPeringkat = hd.indexOf("peringkat");
    var iDaerah = hd.indexOf("daerah");
    var iMp = hd.indexOf("kod_mp");
    var iNama = hd.indexOf("nama_guru");
    var iKp = hd.indexOf("no_kp");
    var iSek = hd.indexOf("kod_sekolah");
    var iUnik = hd.indexOf("no_rujukan_surat");

    for (var i = 1; i < dataKpn.length; i++) {
      if (dataKpn[i][iTahun].toString().trim() == tahun && 
          dataKpn[i][iTahap].toString().trim().toUpperCase() == tahap.toUpperCase() && 
          dataKpn[i][iKp].toString().trim() == noKp) {
        rekod = { 
          tahun: dataKpn[i][iTahun], 
          tahap: dataKpn[i][iTahap],
          peringkat: dataKpn[i][iPeringkat],
          daerah: dataKpn[i][iDaerah],
          kod_mp: dataKpn[i][iMp],
          nama: dataKpn[i][iNama], 
          no_kp: dataKpn[i][iKp], 
          kod_sek: dataKpn[i][iSek],
          id_unik_surat: dataKpn[i][iUnik] 
        }; 
        break;
      }
    }
    
    if (!rekod) return isRawHtml ? "Ralat: Rekod Pentaksir tidak dijumpai." : { status: "FAILED", message: "Rekod tidak dijumpai." };

    // 2. Tarik Maklumat Sekolah Asal
    var sheetSek = ss.getSheetByName('tbsekolah');
    var dataSek = sheetSek ? sheetSek.getDataRange().getValues() : [];
    var namaSek = rekod.kod_sek, poskodSek = "", bandarSek = "";
    for (var s = 1; s < dataSek.length; s++) {
      if (dataSek[s][5] && dataSek[s][5].toString().trim().toUpperCase() === rekod.kod_sek.toUpperCase()) {
         namaSek = dataSek[s][6] || rekod.kod_sek;
         poskodSek = dataSek[s][8] || "";
         bandarSek = dataSek[s][9] || "";
         break;
      }
    }

    // 3. Tarik Maklumat Mata Pelajaran
    var sheetMp = ss.getSheetByName('tbmp');
    var dataMp = sheetMp ? sheetMp.getDataRange().getValues() : [];
    var namaMp = rekod.kod_mp;
    for (var m = 1; m < dataMp.length; m++) {
      if (dataMp[m][1] && dataMp[m][1].toString().trim().toUpperCase() === rekod.kod_mp.toUpperCase()) {
         namaMp = dataMp[m][2] || rekod.kod_mp;
         break;
      }
    }

    // 4. Tarik Maklumat Surat & Sain Pengarah
    var sheetSurat = ss.getSheetByName('tbsurat');
    var dataSurat = sheetSurat ? sheetSurat.getDataRange().getValues() : [];
    var konfigSurat = null;
    var idDicari = (rekod.id_unik_surat || "").toString().trim().toUpperCase();

    for (var j = 1; j < dataSurat.length; j++) {
      var idRow = dataSurat[j][0] ? dataSurat[j][0].toString().trim().toUpperCase() : "";
      if (idDicari !== "" && idRow === idDicari) {
        konfigSurat = { 
          nama_surat: dataSurat[j][1], 
          rujukan_fail: dataSurat[j][2], 
          tarikh_surat: dataSurat[j][3], 
          nama_pengarah: dataSurat[j][4], 
          pingat_pengarah: dataSurat[j][5], 
          jawatan_pengarah: dataSurat[j][6], 
          nama_jabatan: dataSurat[j][7], 
          id_gambar_sain: dataSurat[j][8] 
        }; break;
      }
    }
    
    if (!konfigSurat) return isRawHtml ? "Ralat: Konfigurasi Surat tidak dijumpai." : { status: "FAILED", message: "Tetapan Surat tidak dijumpai dalam rekod tbsurat." };

    // 5. Ganti Kata Kunci pada Template HTML
    var htmlTemplate = HtmlService.createHtmlOutputFromFile('surat_lantikan_kpn_kpkw').getContent();
    
    // SIFU FIX: Logik Tarikh Surat (Guna database jika ada, jika tiada guna auto komputer)
    var bM = ["Januari","Februari","Mac","April","Mei","Jun","Julai","Ogos","September","Oktober","November","Disember"];
    var tK = new Date(); 
    var tarikhAuto = tK.getDate() + " " + bM[tK.getMonth()] + " " + tK.getFullYear();
    var tarikhPaparan = (konfigSurat.tarikh_surat && konfigSurat.tarikh_surat.toString().trim() !== "") ? konfigSurat.tarikh_surat.toString().trim() : tarikhAuto;

    var htmlSainPengarah = (konfigSurat.id_gambar_sain && konfigSurat.id_gambar_sain.trim() !== "") 
      ? '<img src="https://drive.google.com/uc?export=view&id=' + konfigSurat.id_gambar_sain.trim() + '" style="height: 90px; width: auto; margin-top: -15px; margin-bottom: -30px;">' : "t.t";

    var isNegeri = rekod.peringkat.toUpperCase() === "NEGERI";
    
    // PEMBETULAN NAMA JAWATAN & DAERAH
    var jawatanTitle = isNegeri ? "KETUA PENTAKSIR NEGERI" : "KETUA PENTAKSIR KAWASAN";
    var namaDaerah = isNegeri ? "NEGERI PERAK" : getNamaDaerahFullBackend(rekod.daerah, ss);

    // MENGGANTIKAN TEKS BAYARAN DENGAN RM YANG TEPAT
    var bayaranIkhtisasTepat = isNegeri ? "RM75.00" : "RM50.00";

    htmlTemplate = htmlTemplate.replace(/{{RUJUKAN}}/g, konfigSurat.rujukan_fail || "-");
    htmlTemplate = htmlTemplate.replace(/{{NAMA_SURAT}}/g, konfigSurat.nama_surat || "MATRIKS PEMBELAJARAN"); 
    htmlTemplate = htmlTemplate.replace(/{{TARIKH}}/g, tarikhPaparan);
    htmlTemplate = htmlTemplate.replace(/{{NAMA}}/g, rekod.nama || "");
    htmlTemplate = htmlTemplate.replace(/{{NO_KP}}/g, rekod.no_kp || "");
    htmlTemplate = htmlTemplate.replace(/{{SEKOLAH_ASAL}}/g, namaSek || "");
    htmlTemplate = htmlTemplate.replace(/{{POSKOD}}/g, poskodSek || "");
    htmlTemplate = htmlTemplate.replace(/{{BANDAR}}/g, bandarSek || "");
    htmlTemplate = htmlTemplate.replace(/{{TAHUN}}/g, rekod.tahun || "");
    
    htmlTemplate = htmlTemplate.replace(/{{JAWATAN_TITLE}}/g, jawatanTitle);
    htmlTemplate = htmlTemplate.replace(/{{MATA_PELAJARAN}}/g, namaMp);
    htmlTemplate = htmlTemplate.replace(/{{KAWASAN}}/g, namaDaerah);
    
    // Nilai baharu
    htmlTemplate = htmlTemplate.replace(/{{BAYARAN}}/g, bayaranIkhtisasTepat);
    
    htmlTemplate = htmlTemplate.replace(/{{SAIN_PENGARAH}}/g, htmlSainPengarah); 
    htmlTemplate = htmlTemplate.replace(/{{NAMA_PENGARAH}}/g, konfigSurat.nama_pengarah || "");
    htmlTemplate = htmlTemplate.replace(/{{PINGAT_PENGARAH}}/g, konfigSurat.pingat_pengarah ? ", " + konfigSurat.pingat_pengarah : "");
    htmlTemplate = htmlTemplate.replace(/{{JAWATAN_PENGARAH}}/g, konfigSurat.jawatan_pengarah || "");
    htmlTemplate = htmlTemplate.replace(/{{NAMA_JABATAN}}/g, konfigSurat.nama_jabatan || "");
    
    if (isRawHtml) return htmlTemplate;

    // 6. Tukar Kepada Fail PDF (BLOB)
    var blobPDF = Utilities.newBlob(htmlTemplate, MimeType.HTML).getAs(MimeType.PDF);
    var cleanName = (rekod.nama).replace(/[^a-zA-Z0-9 ]/g, "").replace(/\s+/g, "_");
    blobPDF.setName("SuratLantikan_" + (isNegeri ? "KPN" : "KPKw") + "_" + cleanName + "_" + tahun + ".pdf");

    return { 
      status: "SUCCESS", 
      base64: Utilities.base64Encode(blobPDF.getBytes()), 
      filename: blobPDF.getName() 
    };

  } catch (e) { 
    if (isRawHtml) return "Ralat: " + e.message;
    return { status: "FAILED", message: "Ralat Penjanaan PDF HTML: " + e.message }; 
  }
}

// =========================================================================
// FUNGSI PEMBANTU: CARI NAMA DAERAH PENUH DARI PANGKALAN DATA
// =========================================================================
function getNamaDaerahFullBackend(kodAtauPpd, ss) {
    if (!kodAtauPpd) return "-";
    var sheet = ss.getSheetByName("tbdaerah");
    if (!sheet) return kodAtauPpd;
    
    var data = sheet.getDataRange().getValues();
    var inputSemak = kodAtauPpd.toString().trim().toUpperCase();
    
    for(var i=1; i<data.length; i++) {
        var dbKod = data[i][0] ? data[i][0].toString().trim().toUpperCase() : "";
        var dbPpd = data[i][5] ? data[i][5].toString().trim().toUpperCase() : "";
        var dbNama = data[i][2] ? data[i][2].toString().trim().toUpperCase() : "";
        
        if (dbKod === inputSemak || dbPpd === inputSemak || dbNama === inputSemak) {
            return dbNama; // Hanya kembalikan nama penuh sahaja, tanpa kod.
        }
    }
    return kodAtauPpd;
}

// =========================================================================
// FUNGSI BANTUAN LAPORAN: MENDAPATKAN SENARAI MP UNTUK DROPDOWN
// =========================================================================
function dapatkanSenaraiMpLaporanPukal(jenisPeperiksaan) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ambilSheetKalisRalat(ss, 'tbmp');
    var data = sheet.getDataRange().getValues();
    var list = [];
    
    // Tentukan Tahap berdasarkan parameter yang dihantar dari frontend
    var cariTahap = "RENDAH";
    if (jenisPeperiksaan && jenisPeperiksaan.toString().toUpperCase().indexOf("3") > -1) {
      cariTahap = "MENENGAH";
    }
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][1]) {
        var kodMp = data[i][1].toString().trim();
        var namaMp = data[i][2] ? data[i][2].toString().trim() : "";
        var tahapMp = data[i][3] ? data[i][3].toString().trim().toUpperCase() : "UMUM";
        
        // Hanya masukkan subjek yang sama tahap atau UMUM
        if (tahapMp === cariTahap || tahapMp === "UMUM") {
          list.push({ kod: kodMp, nama: namaMp });
        }
      }
    }
    list.sort((a,b) => a.nama.localeCompare(b.nama));
    return list;
  } catch(e) { return []; }
}

// =========================================================================
// TURBO ENGINE: CETAKAN PUKAL SURAT PENYELIA KAWASAN (PK)
// =========================================================================
function janaSuratPukalPK_Turbo(kodDaerah, jenisPeperiksaan, tahun) {
  try {
    var tahapSurat = (jenisPeperiksaan.toUpperCase().indexOf("4") !== -1) ? "TAHUN 4" : "TINGKATAN 3";
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Baca pangkalan data penting sekali sahaja ke dalam memori
    var dataPk = ambilSheetKalisRalat(ss, 'tbpk').getDataRange().getValues();
    var dataSek = ambilSheetKalisRalat(ss, 'tbsekolah').getDataRange().getValues();
    var dataMatrik = ambilSheetKalisRalat(ss, 'tbmatrik_operasi').getDataRange().getValues();
    var dataBk = ambilSheetKalisRalat(ss, 'tbbilikkebal').getDataRange().getValues();
    var dataSurat = ambilSheetKalisRalat(ss, 'tbsurat').getDataRange().getValues();
    
    var mapSekolah = {};
    for (var s = 1; s < dataSek.length; s++) {
      if(dataSek[s][5]) {
        mapSekolah[dataSek[s][5].toString().trim().toUpperCase()] = {
          nama: dataSek[s][6] ? dataSek[s][6].toString().trim() : "",
          poskod: dataSek[s][8] ? dataSek[s][8].toString().trim() : "",
          bandar: dataSek[s][9] ? dataSek[s][9].toString().trim() : ""
        };
      }
    }

    var mapBk = {};
    for (var b = 1; b < dataBk.length; b++) {
      if(dataBk[b][0]) {
        var namaBilik = dataBk[b][1] ? dataBk[b][1].toString().trim() : "";
        var daerahBilik = dataBk[b][2] ? dataBk[b][2].toString().trim().toUpperCase() : "";
        var namaDaerahPenuh = getNamaDaerahFullBackend(daerahBilik, ss);
        mapBk[dataBk[b][0].toString().trim().toUpperCase()] = (namaDaerahPenuh && namaDaerahPenuh !== "-" ? namaDaerahPenuh + " - " : "") + namaBilik;
      }
    }

    var htmlTemplate = HtmlService.createHtmlOutputFromFile('surat_lantikan_pk').getContent();
    
    var bM = ["Januari","Februari","Mac","April","Mei","Jun","Julai","Ogos","September","Oktober","November","Disember"];
    var tK = new Date(); 
    var tarikhAuto = tK.getDate() + " " + bM[tK.getMonth()] + " " + tK.getFullYear();

    var htmlArray = [];
    var totalCetak = 0;

    for (var i = 1; i < dataPk.length; i++) {
      var row = dataPk[i];
      var pTahun = row[1] ? row[1].toString().trim() : "";
      var pDaerah = row[3] ? row[3].toString().trim().toUpperCase() : "";
      var pTahap = row[4] ? row[4].toString().trim().toUpperCase() : "";

      if (pTahun === tahun.toString() && pTahap === tahapSurat && (kodDaerah === "SEMUA" || pDaerah === kodDaerah.toUpperCase())) {
        var noKp = row[8] ? row[8].toString().trim() : "";
        var namaPk = row[7] ? row[7].toString().trim() : "";
        var kodSekAsal = row[2] ? row[2].toString().trim().toUpperCase() : "";
        var idUnik = row[12] ? row[12].toString().trim().toUpperCase() : "";

        // Cari sekolah jagaan & BK
        var senaraiSekJagaan = [];
        var firstBk = "";
        for (var m = 1; m < dataMatrik.length; m++) {
          if (dataMatrik[m][1].toString().trim() === tahun.toString() && dataMatrik[m][4].toString().trim().toUpperCase() === tahapSurat) {
            if (dataMatrik[m][6] == noKp || dataMatrik[m][7] == noKp) {
               var kdJagaan = dataMatrik[m][2] ? dataMatrik[m][2].toString().trim().toUpperCase() : "";
               var nmJagaan = mapSekolah[kdJagaan] ? mapSekolah[kdJagaan].nama : kdJagaan;
               if (senaraiSekJagaan.length < 16) senaraiSekJagaan.push(kdJagaan + " - " + nmJagaan);
               if (firstBk === "") firstBk = dataMatrik[m][5] ? dataMatrik[m][5].toString().trim() : "";
            }
          }
        }
        
        var txtJagaan = "";
        if (senaraiSekJagaan.length > 0) {
          for(var r = 0; r < senaraiSekJagaan.length; r++) {
            txtJagaan += "<tr><td style='text-align: center; font-weight: bold;'>" + (r + 1) + "</td><td><b>" + senaraiSekJagaan[r] + "</b></td></tr>";
          }
        } else {
          txtJagaan = "<tr><td colspan='2' style='text-align: center; padding: 15px;'>TIADA REKOD SEKOLAH JAGAAN DITETAPKAN</td></tr>";
        }
        
        var txtBk = firstBk ? (mapBk[firstBk.toUpperCase()] ? (mapBk[firstBk.toUpperCase()]) : firstBk) : "BELUM DITETAPKAN";

        var conf = null;
        for (var c = 1; c < dataSurat.length; c++) {
          var idR = dataSurat[c][0] ? dataSurat[c][0].toString().trim().toUpperCase() : "";
          if (idUnik !== "" && idR === idUnik) { conf = dataSurat[c]; break; }
        }
        if (!conf) {
          for (var c = 1; c < dataSurat.length; c++) {
            var idR = dataSurat[c][0] ? dataSurat[c][0].toString().trim().toUpperCase() : "";
            if (idR.indexOf("PK") > -1 && idR.indexOf(tahapSurat.replace(/\s+/g,'')) > -1) { conf = dataSurat[c]; break; }
          }
        }

        var sNama = mapSekolah[kodSekAsal] ? mapSekolah[kodSekAsal].nama : kodSekAsal;
        var sPoskod = mapSekolah[kodSekAsal] ? mapSekolah[kodSekAsal].poskod : "";
        var sBandar = mapSekolah[kodSekAsal] ? mapSekolah[kodSekAsal].bandar : "";
        
        var imgSain = (conf && conf[8] && conf[8].toString().trim() !== "") ? '<img src="https://drive.google.com/uc?export=view&id=' + conf[8].toString().trim() + '" style="height: 90px; width: auto; margin-top: -15px; margin-bottom: -30px;">' : "t.t";
        var tarikhPaparan = (conf && conf[3] && conf[3].toString().trim() !== "") ? conf[3].toString().trim() : tarikhAuto;

        var htmlIndividu = htmlTemplate
          .replace(/{{RUJUKAN}}/g, conf ? conf[2] : "-") 
          .replace(/{{NAMA_SURAT}}/g, conf ? conf[1] : "MATRIKS PEMBELAJARAN") 
          .replace(/{{TARIKH}}/g, tarikhPaparan)
          .replace(/{{NAMA}}/g, (namaPk).toUpperCase())
          .replace(/{{NO_KP}}/g, noKp)
          .replace(/{{SEKOLAH_ASAL}}/g, sNama.toUpperCase())
          .replace(/{{POSKOD}}/g, sPoskod)
          .replace(/{{BANDAR}}/g, sBandar.toUpperCase())
          .replace(/{{TAHUN}}/g, tahun)
          .replace(/{{JAWATAN_BESAR}}/g, "PENYELIA KAWASAN") 
          .replace(/{{JAWATAN_TITLE}}/g, "Penyelia Kawasan")
          .replace(/{{TARIKH_UJIAN}}/g, conf ? conf[9] : "AKAN DIMAKLUMKAN")
          .replace(/{{TR_SEKOLAH_JAGAAN}}/g, txtJagaan.toUpperCase())
          .replace(/{{BILIK_KEBAL}}/g, txtBk.toUpperCase())
          .replace(/{{SAIN_PENGARAH}}/g, imgSain)
          .replace(/{{NAMA_PENGARAH}}/g, conf ? conf[4].toString().toUpperCase() : "")
          .replace(/{{PINGAT_PENGARAH}}/g, (conf && conf[5]) ? (", " + conf[5]) : "")
          .replace(/{{JAWATAN_PENGARAH}}/g, conf ? conf[6] : "")
          .replace(/{{NAMA_JABATAN}}/g, conf ? conf[7] : "");

        htmlArray.push(htmlIndividu);
        totalCetak++;
      }
    }

    if (totalCetak === 0) return { status: "FAILED", message: "Tiada rekod Penyelia Kawasan (PK) ditemui bagi tapisan yang dipilih." };

    var finalHtml = htmlArray.join("");
    
    var blobPDF = Utilities.newBlob(finalHtml, MimeType.HTML).getAs(MimeType.PDF);
    blobPDF.setName("Surat_PK_Pukal_" + kodDaerah + "_" + tahun + ".pdf");
    return { status: "SUCCESS", base64: Utilities.base64Encode(blobPDF.getBytes()), filename: blobPDF.getName() };
  } catch(e) { return { status: "FAILED", message: e.message }; }
}

// =========================================================================
// TURBO ENGINE: CETAKAN PUKAL SURAT PENTAKSIR (KPN / KPKw)
// =========================================================================
function janaSuratPukalPentaksir_Turbo(kategori, kodMp, kodDaerah, jenisPeperiksaan, tahun) {
  try {
    var tahapSurat = (jenisPeperiksaan.toUpperCase().indexOf("4") !== -1) ? "TAHUN 4" : "TINGKATAN 3";
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var dataKpn = ambilSheetKalisRalat(ss, 'tbkpn_kpkw').getDataRange().getValues();
    var dataSek = ambilSheetKalisRalat(ss, 'tbsekolah').getDataRange().getValues();
    var dataSurat = ambilSheetKalisRalat(ss, 'tbsurat').getDataRange().getValues();
    
    var mapSekolah = {};
    for (var s = 1; s < dataSek.length; s++) {
      if(dataSek[s][5]) mapSekolah[dataSek[s][5].toString().trim().toUpperCase()] = { nama: dataSek[s][6] ? dataSek[s][6].toString().trim() : "", poskod: dataSek[s][8] ? dataSek[s][8].toString().trim() : "", bandar: dataSek[s][9] ? dataSek[s][9].toString().trim() : "" };
    }
    
    var mapMp = {};
    var dataTbMp = ambilSheetKalisRalat(ss, 'tbmp').getDataRange().getValues();
    for (var m = 1; m < dataTbMp.length; m++) {
      if(dataTbMp[m][1]) mapMp[dataTbMp[m][1].toString().trim().toUpperCase()] = dataTbMp[m][2] ? dataTbMp[m][2].toString().trim() : "";
    }

    var htmlTemplate = HtmlService.createHtmlOutputFromFile('surat_lantikan_kpn_kpkw').getContent();
    
    var bM = ["Januari","Februari","Mac","April","Mei","Jun","Julai","Ogos","September","Oktober","November","Disember"];
    var tK = new Date(); 
    var tarikhAuto = tK.getDate() + " " + bM[tK.getMonth()] + " " + tK.getFullYear();

    var htmlKeseluruhan = "";
    var totalCetak = 0;

    var peringkatTarget = (kategori === "KPKW") ? "DAERAH" : "NEGERI";

    for (var i = 1; i < dataKpn.length; i++) {
      var row = dataKpn[i];
      var rTahun = row[1] ? row[1].toString().trim() : "";
      var rTahap = row[2] ? row[2].toString().trim().toUpperCase() : "";
      var rPeringkat = row[3] ? row[3].toString().trim().toUpperCase() : "";
      var rDaerah = row[4] ? row[4].toString().trim().toUpperCase() : "";
      var rMp = row[5] ? row[5].toString().trim().toUpperCase() : "";

      var matchTahap = (rTahun === tahun.toString() && rTahap === tahapSurat && rPeringkat === peringkatTarget);
      var matchMp = (kodMp === "SEMUA" || rMp === kodMp.toUpperCase());
      var matchDaerah = (kategori === "KPN") ? true : (kodDaerah === "SEMUA" || rDaerah === kodDaerah.toUpperCase());

      if (matchTahap && matchMp && matchDaerah) {
        var namaGuru = row[6] ? row[6].toString().trim() : "";
        var noKp = row[7] ? row[7].toString().trim() : "";
        var kodSek = row[8] ? row[8].toString().trim().toUpperCase() : "";
        var idUnik = row[11] ? row[11].toString().trim().toUpperCase() : "";

        var conf = null;
        for (var c = 1; c < dataSurat.length; c++) {
          var idR = dataSurat[c][0] ? dataSurat[c][0].toString().trim().toUpperCase() : "";
          if (idUnik !== "" && idR === idUnik) { conf = dataSurat[c]; break; }
        }
        if (!conf) {
          for (var c = 1; c < dataSurat.length; c++) {
             var idR = dataSurat[c][0] ? dataSurat[c][0].toString().trim().toUpperCase() : "";
             var cariPrefix = kategori === "KPN" ? "KPN" : "KPKW";
             if (idR.indexOf(cariPrefix) > -1 && idR.indexOf(tahapSurat.replace(/\s+/g,'')) > -1) { conf = dataSurat[c]; break; }
          }
        }

        var sNama = mapSekolah[kodSek] ? mapSekolah[kodSek].nama : kodSek;
        var sPoskod = mapSekolah[kodSek] ? mapSekolah[kodSek].poskod : "";
        var sBandar = mapSekolah[kodSek] ? mapSekolah[kodSek].bandar : "";
        var nmMpFull = mapMp[rMp] ? mapMp[rMp] : rMp;
        var namaDaeKawasan = kategori === "KPN" ? "NEGERI PERAK" : getNamaDaerahFullBackend(rDaerah, ss);
        
        var imgSain = (conf && conf[8] && conf[8].toString().trim() !== "") ? '<img src="https://drive.google.com/uc?export=view&id=' + conf[8].toString().trim() + '" style="height: 90px; width: auto; margin-top: -15px; margin-bottom: -30px;">' : "t.t";

        var bayaranIkhtisasTepat = (kategori === "KPN") ? "RM75.00" : "RM50.00";

        var tarikhPaparan = (conf && conf[3] && conf[3].toString().trim() !== "") ? conf[3].toString().trim() : tarikhAuto;

        var htmlIndividu = htmlTemplate
          .replace(/{{RUJUKAN}}/g, conf ? conf[2] : "-") 
          .replace(/{{NAMA_SURAT}}/g, conf ? conf[1] : "MATRIKS PEMBELAJARAN") 
          .replace(/{{TARIKH}}/g, tarikhPaparan)
          .replace(/{{NAMA}}/g, (namaGuru).toUpperCase())
          .replace(/{{NO_KP}}/g, noKp)
          .replace(/{{SEKOLAH_ASAL}}/g, sNama.toUpperCase())
          .replace(/{{POSKOD}}/g, sPoskod)
          .replace(/{{BANDAR}}/g, sBandar.toUpperCase())
          .replace(/{{TAHUN}}/g, tahun)
          .replace(/{{JAWATAN_TITLE}}/g, kategori === "KPN" ? "KETUA PENTAKSIR NEGERI" : "KETUA PENTAKSIR KAWASAN")
          .replace(/{{MATA_PELAJARAN}}/g, nmMpFull.toUpperCase())
          .replace(/{{KAWASAN}}/g, namaDaeKawasan.toUpperCase())
          .replace(/{{BAYARAN}}/g, bayaranIkhtisasTepat)
          .replace(/{{SAIN_PENGARAH}}/g, imgSain)
          .replace(/{{NAMA_PENGARAH}}/g, conf ? conf[4].toString().toUpperCase() : "")
          .replace(/{{PINGAT_PENGARAH}}/g, (conf && conf[5]) ? (", " + conf[5]) : "")
          .replace(/{{JAWATAN_PENGARAH}}/g, conf ? conf[6] : "")
          .replace(/{{NAMA_JABATAN}}/g, conf ? conf[7] : "");

        htmlKeseluruhan += htmlIndividu + "<div style='page-break-after: always;'></div>";
        totalCetak++;
      }
    }

    if (totalCetak === 0) return { status: "FAILED", message: `Tiada rekod ${kategori} ditemui bagi tapisan yang dipilih.` };

    var blobPDF = Utilities.newBlob(htmlKeseluruhan, MimeType.HTML).getAs(MimeType.PDF);
    blobPDF.setName(`Surat_${kategori}_Pukal_${tahun}.pdf`);
    return { status: "SUCCESS", base64: Utilities.base64Encode(blobPDF.getBytes()), filename: blobPDF.getName() };
  } catch(e) { return { status: "FAILED", message: e.message }; }
}

// =========================================================================
// SALIN PUKAL KPN / KPKw (KEBAL ID & ANTI-PENDUA)
// =========================================================================
function salinPukalKpnKpkwBackend(senaraiIdString, targetTahun, tahap) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); 

    var senaraiIdCari = JSON.parse(senaraiIdString);
    if (senaraiIdCari.length === 0) {
      lock.releaseLock();
      return { status: "WARNING", message: "Tiada rekod dipilih." };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ambilSheetKalisRalat(ss, 'tbkpn_kpkw');
    var data = sheet.getDataRange().getValues();

    var existingKeys = new Set();
    var maxId = 0;

    for (var i = 1; i < data.length; i++) {
      var rowId = parseInt(data[i][0]) || 0;
      if (rowId > maxId) maxId = rowId; 

      var dbTahun = data[i][1] ? data[i][1].toString().trim() : "";
      var dbTahap = data[i][2] ? data[i][2].toString().trim().toUpperCase() : "";
      var dbPeringkat = data[i][3] ? data[i][3].toString().trim().toUpperCase() : "";
      var dbDaerah = data[i][4] ? data[i][4].toString().trim().toUpperCase() : "";
      var dbMp = data[i][5] ? data[i][5].toString().trim().toUpperCase() : "";
      var dbNoKp = data[i][7] ? data[i][7].toString().trim() : "";

      if (dbTahun === targetTahun) {
        var uniqueKey = dbTahap + "_" + dbPeringkat + "_" + dbDaerah + "_" + dbMp;
        existingKeys.add(uniqueKey); 
        if (dbNoKp !== "") existingKeys.add(dbNoKp); 
      }
    }

    var copiedCount = 0;
    var duplicateCount = 0;
    var newRows = [];

    for (var j = 1; j < data.length; j++) {
      var rowIdStr = data[j][0] ? data[j][0].toString() : "";
      
      if (senaraiIdCari.includes(rowIdStr)) {
        var cTahap = data[j][2] ? data[j][2].toString().trim().toUpperCase() : "";
        var cPeringkat = data[j][3] ? data[j][3].toString().trim().toUpperCase() : "";
        var cDaerah = data[j][4] ? data[j][4].toString().trim().toUpperCase() : "";
        var cMp = data[j][5] ? data[j][5].toString().trim().toUpperCase() : "";
        var cNama = data[j][6] ? data[j][6].toString().toUpperCase() : "";
        var cNoKp = data[j][7] ? data[j][7].toString().trim() : "";
        var cSek = data[j][8] ? data[j][8].toString().toUpperCase() : "";
        var cTel = data[j][9] ? data[j][9].toString() : "";
        var cEmel = data[j][10] ? data[j][10].toString() : "";

        var checkKey = cTahap + "_" + cPeringkat + "_" + cDaerah + "_" + cMp;

        if (existingKeys.has(checkKey) || (cNoKp !== "" && existingKeys.has(cNoKp))) {
          duplicateCount++;
        } else {
          maxId++; 
          
          var prefixSurat = cPeringkat === "NEGERI" ? "KPN" : "KPKw";
          var idMatrikTahap = (cTahap === "TAHUN 4") ? "MP4" : "MP3";
          var idUnikSurat = targetTahun + prefixSurat + cMp + idMatrikTahap;

          newRows.push([
            maxId, 
            targetTahun,
            cTahap,
            cPeringkat,
            cDaerah,
            cMp,
            cNama,
            "'" + cNoKp,
            cSek,
            cTel ? "'" + cTel : "",
            cEmel,
            idUnikSurat
          ]);
          
          existingKeys.add(checkKey);
          if (cNoKp !== "") existingKeys.add(cNoKp);
          copiedCount++;
        }
      }
    }

    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
      SpreadsheetApp.flush();
    }

    lock.releaseLock();

    if (copiedCount > 0 && duplicateCount === 0) {
      return { status: "SUCCESS", message: copiedCount + " rekod berjaya disalin ke Sesi " + targetTahun + "." };
    } else if (copiedCount > 0 && duplicateCount > 0) {
      return { status: "WARNING", message: copiedCount + " rekod berjaya disalin. " + duplicateCount + " rekod diabaikan kerana pertindihan guru/jawatan pada Sesi " + targetTahun + "." };
    } else {
      return { status: "WARNING", message: "Tiada rekod disalin. Semua rekod yang dipilih telah wujud pada Sesi " + targetTahun + "." };
    }

  } catch (e) {
    if (lock) lock.releaseLock();
    return { status: "FAILED", message: e.message };
  }
}

// =========================================================================
// MODUL OPERASI SEKOLAH (PENGAWAS & CADANGAN KP) - KEMASKINI 0 MURID
// =========================================================================
function getDataOperasiSekolahUtama(kodSekolah, tahun) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var kodSek = kodSekolah.toString().trim().toUpperCase();
    var data = {
      jum_calon: 0, kategori: "RENDAH", tahap_matrik: "TAHUN 4", daerah: "",
      profilSekolah: null,
      senaraiKp: [], senaraiPgws: [], senaraiPusat: [],
      tetapan: {} 
    };

    try {
      var sheetTetapan = ss.getSheetByName('tb_tetapan');
      if (sheetTetapan) {
        var dataTetapan = sheetTetapan.getDataRange().getValues();
        for (var t = 1; t < dataTetapan.length; t++) {
          data.tetapan[dataTetapan[t][0]] = dataTetapan[t][1];
        }
      }
    } catch(e) {}
    
    var sheetSek = ambilSheetKalisRalat(ss, 'tbsekolah');
    var dataSek = sheetSek.getDataRange().getValues();
    for(var s=1; s<dataSek.length; s++) {
      if(dataSek[s][5] && dataSek[s][5].toString().trim().toUpperCase() === kodSek) {
        data.profilSekolah = {
          rowNum: s + 1, 
          kod: dataSek[s][5] ? dataSek[s][5].toString().trim().toUpperCase() : "",
          ppd: dataSek[s][3] ? dataSek[s][3].toString().trim() : "",
          nama_gb: dataSek[s][13] ? dataSek[s][13].toString().trim() : "",
          kp_gb: dataSek[s][14] ? dataSek[s][14].toString().trim() : "",
          tel_gb: dataSek[s][15] ? dataSek[s][15].toString().trim() : "",
          emel_gb: dataSek[s][16] ? dataSek[s][16].toString().trim() : "",
          pengesahan_pengawas: dataSek[s][25] ? dataSek[s][25].toString().trim() : "",
          status_lokasi: dataSek[s][26] ? dataSek[s][26].toString().trim().toUpperCase() : "" // TANGKAP LOKASI
        };
        data.jum_calon = parseInt(dataSek[s][23]) || 0;
        data.kategori = dataSek[s][1] ? dataSek[s][1].toString().trim().toUpperCase() : "RENDAH";
        data.daerah = dataSek[s][3] ? dataSek[s][3].toString().trim().toUpperCase() : "";
        data.tahap_matrik = (data.kategori === "MENENGAH") ? "TINGKATAN 3" : "TAHUN 4";
        break;
      }
    }
    
    var sheetKp = ss.getSheetByName('tbkp');
    if (sheetKp) {
      var dKp = sheetKp.getDataRange().getValues();
      for(var k=1; k<dKp.length; k++) {
        if(dKp[k][1] == tahun && dKp[k][2].toString().trim().toUpperCase() === kodSek && dKp[k][4].toString().trim().toUpperCase() === data.tahap_matrik) {
          var kodCdgn = dKp[k][11] ? dKp[k][11].toString().trim().toUpperCase() : "";
          var namaCdgn = "";
          for(var sc=1; sc<dataSek.length; sc++) {
            if(dataSek[sc][5] && dataSek[sc][5].toString().trim().toUpperCase() === kodCdgn) {
              namaCdgn = dataSek[sc][6] || ""; break;
            }
          }
          data.senaraiKp.push({
            rowNum: k+1, 
            matrik_tahap: dKp[k][4], 
            nama_guru: dKp[k][6], 
            no_kp: dKp[k][7], 
            no_tel: dKp[k][8], 
            emel: dKp[k][9], 
            alamat_rumah: dKp[k][10],
            kod_cadangan: kodCdgn, 
            paparan_cadangan: kodCdgn + (namaCdgn ? " - " + namaCdgn : ""), 
            status: dKp[k][13]
          });
        }
      }
    }
    
    var sheetPgws = ss.getSheetByName('tbp');
    if (!sheetPgws) {
      sheetPgws = ss.insertSheet('tbp');
      sheetPgws.appendRow(['ID', 'TAHUN', 'KOD_SEKOLAH', 'DAERAH', 'MATRIK_TAHAP', 'NAMA_PENGAWAS', 'NO_KP', 'JANTINA', 'NO_TEL', 'EMEL']);
    }
    var dPgws = sheetPgws.getDataRange().getValues();
    for(var p=1; p<dPgws.length; p++) {
      if(dPgws[p][1] == tahun && dPgws[p][2].toString().trim().toUpperCase() === kodSek && dPgws[p][4].toString().trim().toUpperCase() === data.tahap_matrik) {
        data.senaraiPgws.push({
          rowNum: p+1, id: dPgws[p][0], nama: dPgws[p][5], no_kp: dPgws[p][6], jantina: dPgws[p][7], no_tel: dPgws[p][8], emel: dPgws[p][9]
        });
      }
    }
    
    // TANGKAP SENARAI PUSAT
    data.senaraiPusat = dapatkanSekolahDaerahSama(data.daerah, kodSek);
    
    return { status: "SUCCESS", data: data };
  } catch(e) { return { status: "FAILED", message: e.message }; }
}

// =========================================================================
// SIMPAN MAKLUMAT PENGAWAS BARU (tbp)
// =========================================================================
function simpanPengawasBackend(d) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('tbp');
    
    if (!sheet) {
      sheet = ss.insertSheet('tbp');
      sheet.appendRow(['ID', 'TAHUN', 'KOD_SEKOLAH', 'DAERAH', 'MATRIK_TAHAP', 'NAMA_PENGAWAS', 'NO_KP', 'JANTINA', 'NO_TEL', 'EMEL']);
    }
    
    var r = d.rowNum ? Number(d.rowNum) : 0;
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (r > 0 && r === (i + 1)) continue;
      
      var dbTahun = data[i][1] ? data[i][1].toString().trim() : "";
      var dbKodSek = data[i][2] ? data[i][2].toString().trim().toUpperCase() : "";
      var dbKp = data[i][6] ? data[i][6].toString().trim() : "";
      
      var inputKodSek = d.kod_sekolah ? d.kod_sekolah.toString().trim().toUpperCase() : "";
      
      if (dbTahun === d.tahun.toString() && dbKodSek === inputKodSek && dbKp === d.no_kp) {
        return { status: "FAILED", message: "No. Kad Pengenalan ini telah didaftarkan sebagai pengawas untuk sesi ini!" };
      }
    }
    
    if (r > 0) {
      sheet.getRange(r, 6).setValue(d.nama.toUpperCase());
      sheet.getRange(r, 7).setValue(d.no_kp);
      sheet.getRange(r, 8).setValue(d.jantina);
      sheet.getRange(r, 9).setValue(d.no_tel);
      sheet.getRange(r, 10).setValue(d.emel);
    } else {
      var nextId = data.length > 1 ? Number(data[data.length - 1][0]) + 1 : 1;
      var inputKod = d.kod_sekolah ? d.kod_sekolah.toString().toUpperCase() : "";
      sheet.appendRow([nextId, d.tahun, inputKod, d.daerah, d.tahap, d.nama.toUpperCase(), d.no_kp, d.jantina, d.no_tel, d.emel]);
    }
    
    SpreadsheetApp.flush();
    return { status: "SUCCESS" };
  } catch(e) { 
    return { status: "FAILED", message: e.message }; 
  }
}

// ==========================================================
// KEMASKINI STATUS PENGESAHAN OPERASI PENGAWAS
// ==========================================================
function kemaskiniStatusPengesahanOperasiDB(rowDB, status) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('tbsekolah');
    sheet.getRange(rowDB, 26).setValue(status);
    return {status: "SUCCESS"};
  } catch(e) {
    return {status: "FAILED", message: e.message};
  }
}

// ==========================================================
// 20. FUNGSI KHAS: JANA BORANG SPP2 UNTUK 1 PUSAT SAHAJA
// ==========================================================
function janaBorangSPP2PDF_Pusat(kodPusatCari, jenisPeperiksaan, tahun) {
  try {
    jenisPeperiksaan = jenisPeperiksaan || "MATRIKS PEMBELAJARAN TAHUN 4";
    tahun = tahun || new Date().getFullYear().toString(); 
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var wsSekolahMaster = ss.getSheetByName("tbsekolah");
    var mapSekolah = {}; var mapCalon = {};
    if (wsSekolahMaster) {
       var dataSek = wsSekolahMaster.getDataRange().getValues();
       for(var s = 1; s < dataSek.length; s++){
          var kodS = dataSek[s][5] ? dataSek[s][5].toString().trim().toUpperCase() : "";
          if(kodS) { mapSekolah[kodS] = dataSek[s][6] ? dataSek[s][6].toString().trim().toUpperCase() : ""; mapCalon[kodS] = dataSek[s][23] ? dataSek[s][23].toString().trim() : "0"; }
       }
    }

    var wsDaerah = ss.getSheetByName("tbdaerah");
    var mapDaerah = {};
    if (wsDaerah) {
      var dataDae = wsDaerah.getDataRange().getValues();
      for(var d = 1; d < dataDae.length; d++){
         var kodD = dataDae[d][5] ? dataDae[d][5].toString().trim().toUpperCase() : "";
         if(kodD) mapDaerah[kodD] = dataDae[d][2] ? dataDae[d][2].toString().trim().toUpperCase() : "";
      }
    }

    var wsBK = ss.getSheetByName("tbbilikkebal");
    var mapBK = {};
    if (wsBK) {
      var dataBK = wsBK.getDataRange().getValues();
      for(var b = 1; b < dataBK.length; b++){
         var idBK = dataBK[b][0] ? dataBK[b][0].toString().trim().toUpperCase() : ""; 
         if(idBK) mapBK[idBK] = dataBK[b][1] ? dataBK[b][1].toString().trim().toUpperCase() : "";
      }
    }

    var wsMatrik = ss.getSheetByName("tbmatrik_operasi"); var dataMatrik = wsMatrik ? wsMatrik.getDataRange().getValues() : []; if (dataMatrik.length > 0) dataMatrik.shift(); 
    var wsKP = ss.getSheetByName("tbkp"); var dataKP = wsKP ? wsKP.getDataRange().getValues() : []; if (dataKP.length > 0) dataKP.shift();
    var wsPK = ss.getSheetByName("tbpk"); var dataPK = wsPK ? wsPK.getDataRange().getValues() : []; if (dataPK.length > 0) dataPK.shift();
    var wsPgws = ss.getSheetByName("tbp"); var dataPgws = wsPgws ? wsPgws.getDataRange().getValues() : []; if (dataPgws.length > 0) dataPgws.shift();

    var pusatTapis = dataMatrik.filter(function(row) {
      var t = row[1] ? row[1].toString().trim() : "";
      var k = row[2] ? row[2].toString().trim().toUpperCase() : "";
      var m = row[4] ? row[4].toString().toUpperCase().trim() : "";
      var pTahap = (jenisPeperiksaan.toUpperCase().indexOf("3") !== -1) ? (m.indexOf("3") !== -1) : (m.indexOf("4") !== -1);
      return (k === kodPusatCari.toUpperCase()) && (t === tahun.toString().trim()) && pTahap;
    });

    if (pusatTapis.length === 0) return { success: false, message: "Pusat ini belum mempunyai penetapan jadual Operasi Matriks yang direkodkan." };

    var tarikhCetakan = Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "dd/MM/yyyy");
    var htmlCardsArray = [];

    pusatTapis.forEach(function(pusat) {
      var kodPusat = pusat[2] ? pusat[2].toString().trim().toUpperCase() : "";       
      var namaPusatFull = mapSekolah[kodPusat] ? (kodPusat + " - " + mapSekolah[kodPusat]) : kodPusat;
      var kodDaerahRow = pusat[3] ? pusat[3].toString().trim().toUpperCase() : "";
      var namaDaerahFull = mapDaerah[kodDaerahRow] ? (kodDaerahRow + " - " + mapDaerah[kodDaerahRow]) : kodDaerahRow;     
      var idBKRow = pusat[5] ? pusat[5].toString().trim().toUpperCase() : ""; 
      var infoBilikKebal = (idBKRow && idBKRow !== "-") ? (mapBK[idBKRow] ? (idBKRow + " - " + mapBK[idBKRow]) : idBKRow) : "-";
      var jumlahCalon = mapCalon[kodPusat] || "0";
      
      var icPk1 = pusat[6] ? pusat[6].toString().trim() : "";
      var icPk2 = pusat[7] ? pusat[7].toString().trim() : "";
      var senaraiPetugasPusat = [];

      [icPk1, icPk2].forEach(function(ic) {
        if(ic !== "") {
          var pkFound = dataPK.find(function(r) { return r[8].toString().trim() === ic && r[1].toString() === tahun.toString(); });
          if(pkFound) senaraiPetugasPusat.push({ jawatan: "PENYELIA KAWASAN", nama: pkFound[7] || "", sekolahAsal: mapSekolah[pkFound[2]] ? (pkFound[2] + " - " + mapSekolah[pkFound[2]]) : pkFound[2], noKp: pkFound[8] || "", noTel: pkFound[9] || "" });
        }
      });

      dataKP.forEach(function(kp) {
        if (kp[1].toString().trim() === tahun.toString() && kp[12].toString().trim().toUpperCase() === kodPusat && kp[13].toString().trim().toUpperCase() === 'DISAHKAN') {
          var tKp = kp[4] ? kp[4].toString().trim().toUpperCase() : "";
          if ((jenisPeperiksaan.indexOf("3") !== -1 && tKp.indexOf("3") !== -1) || (jenisPeperiksaan.indexOf("4") !== -1 && tKp.indexOf("4") !== -1)) {
            var kSekKp = kp[2] ? kp[2].toString().trim().toUpperCase() : "";
            senaraiPetugasPusat.push({ jawatan: "KETUA PENGAWAS", nama: kp[6] || "", sekolahAsal: mapSekolah[kSekKp] ? (kSekKp + " - " + mapSekolah[kSekKp]) : kSekKp, noKp: kp[7] || "", noTel: kp[8] || "" });
          }
        }
      });

      dataPgws.forEach(function(pgws) {
        if (pgws[1].toString().trim() === tahun.toString() && pgws[2].toString().trim().toUpperCase() === kodPusat) {
          var tPgws = pgws[4] ? pgws[4].toString().trim().toUpperCase() : "";
          if ((jenisPeperiksaan.toUpperCase().indexOf("3") !== -1 && tPgws.indexOf("3") !== -1) || (jenisPeperiksaan.toUpperCase().indexOf("4") !== -1 && tPgws.indexOf("4") !== -1)) {
            var kSekPgws = pgws[2] ? pgws[2].toString().trim().toUpperCase() : "";
            senaraiPetugasPusat.push({ jawatan: "PENGAWAS", nama: pgws[5] || "", sekolahAsal: mapSekolah[kSekPgws] ? (kSekPgws + " - " + mapSekolah[kSekPgws]) : kSekPgws, noKp: pgws[6] || "", noTel: pgws[8] || "" });
          }
        }
      });

      var rowsHTML = "";
      if (senaraiPetugasPusat.length > 0) {
        senaraiPetugasPusat.forEach(function(p, idx) {
          var jantina = tentukanJantinaDariKP(p.noKp);
          rowsHTML += `<tr><td class="text-center">${idx + 1}</td><td>${p.jawatan}</td><td>${p.nama}</td><td>${p.sekolahAsal}</td><td class="text-center">${jantina}</td><td class="text-center" style="white-space: nowrap;">${p.noTel}</td></tr>`;
        });
      } else { rowsHTML = `<tr><td colspan="6" class="text-center" style="padding:15px; color:#666;">Tiada Petugas Didaftarkan Bagi Pusat Ini</td></tr>`; }

      var cardHTML = `
        <div class="report-card">
          <div>
            <div class="top-header-bar">
              <div>KOD PUSAT PEPERIKSAAN: <span>${kodPusat}</span></div>
              <div>BORANG SPP2</div>
            </div>
            <div class="header-center">
              <div class="divider">|</div><div class="title-line">LEMBAGA PEPERIKSAAN</div><div class="title-line">KEMENTERIAN PENDIDIKAN</div>
              <div class="sub-line">PEPERIKSAAN : ${jenisPeperiksaan.toUpperCase()} &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; TAHUN : <span>${tahun}</span></div>
              <div class="sub-line">SENARAI PETUGAS YANG DILANTIK</div>
            </div>
            <div class="info-container">
              <table class="info-table"><tr><td style="width: 275px;">KOD DAN NAMA PUSAT PEPERIKSAAN</td><td style="width: 18px;">:</td><td><span>${namaPusatFull}</span></td></tr></table>
              <div style="margin-top: 15px;">
                <table class="info-table">
                  <tr><td><span class="underline-text">INFO CALON</span></td><td style="text-align: right;">DAERAH : <span>${namaDaerahFull}</span></td></tr>
                  <tr><td>JUMLAH CALON : <span>${jumlahCalon}</span> ORANG</td><td style="text-align: right;">BILIK KEBAL : <span>${infoBilikKebal}</span></td></tr>
                </table>
              </div>
            </div>
            <table class="data-table">
              <thead><tr><th style="width: 4%;">BIL</th><th style="width: 20%;">JAWATAN PETUGAS</th><th style="width: 26%;">NAMA PETUGAS</th><th style="width: 26%;">NAMA SEKOLAH PETUGAS</th><th style="width: 9%;">JANTINA</th><th style="width: 15%; white-space: nowrap;">NO TELEFON</th></tr></thead>
              <tbody>${rowsHTML}</tbody>
            </table>
          </div>
          <div class="footer-section">TARIKH <span class="underline-text">CETAKAN</span> : <span>${tarikhCetakan}</span></div>
        </div>`;
      htmlCardsArray.push(cardHTML);
    });

    var htmlTemplate = HtmlService.createTemplateFromFile('TemplateSPP2');
    htmlTemplate.htmlCards = htmlCardsArray.join('');
    var blob = Utilities.newBlob(htmlTemplate.evaluate().getContent(), MimeType.HTML).getAs(MimeType.PDF);
    blob.setName(`Borang_SPP2_${kodPusatCari}_${tahun}.pdf`);

    return { success: true, pdfBase64: Utilities.base64Encode(blob.getBytes()), filename: blob.getName() };
  } catch (err) { return { success: false, message: err.toString() }; }
}

function dapatkanDataAksesSekolah(daerahPengguna, perananPengguna) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var sheetSekolah = ss.getSheetByName("tbsekolah");
    var dataSekolah = sheetSekolah.getDataRange().getValues();
    var mapNamaSekolah = {}; 
    
    for (var j = 1; j < dataSekolah.length; j++) {
      var kodSek = dataSekolah[j][5]; 
      var namaSek = dataSekolah[j][6]; 
      if (kodSek) {
        mapNamaSekolah[kodSek.toString().toUpperCase()] = namaSek;
      }
    }

    var sheetUser = ss.getSheetByName("user");
    var dataUser = sheetUser.getDataRange().getValues();
    var senarai = [];
    
    for (var i = 1; i < dataUser.length; i++) {
      var baris = dataUser[i];
      var username = baris[1];    
      var password = baris[2];    
      var emel = baris[3];        
      var daerah = baris[4];      
      var role = baris[6];        
      var statusTukar = baris[8]; 
      
      if (role && role.toUpperCase() === "SEKOLAH") {
        if (perananPengguna === "JPN" || perananPengguna === "ADMIN" || daerahPengguna === daerah) {
          var carianKod = username.toString().toUpperCase();
          var namaDitemui = mapNamaSekolah[carianKod] || "TIADA MAKLUMAT NAMA";
          
          senarai.push({
            id_sekolah: username,
            nama_sekolah: namaDitemui,
            katalaluan: password,
            daerah: daerah,
            status_tukar: statusTukar,
            emel: emel 
          });
        }
      }
    }
    
    return { status: "SUCCESS", data: senarai };
    
  } catch (error) {
    return { status: "ERROR", message: error.message };
  }
}

// =========================================================================
// DAPATKAN SENARAI SEMUA MATA PELAJARAN UNTUK DROPDOWN KPMP
// =========================================================================
function dapatkanSenaraiSemuaMp() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetMp = ambilSheetKalisRalat(ss, 'tbmp');
    var data = sheetMp.getDataRange().getValues();
    var senarai = [];
    for (var i = 1; i < data.length; i++) {
      var kod = data[i][1] ? data[i][1].toString().trim() : "";
      var nama = data[i][2] ? data[i][2].toString().trim() : "";
      if (kod) senarai.push({ kod: kod, nama: nama });
    }
    return senarai;
  } catch(e) { return []; }
}

// =========================================================================
// FUNGSI EKSPORT EXCEL / CSV UNTUK KPMP
// =========================================================================
function muatTurunDataKPMP(kodDaerah, kodMp) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var mapSekolah = {};
    var dataSek = ambilSheetKalisRalat(ss, 'tbsekolah').getDataRange().getValues();
    for(var s = 1; s < dataSek.length; s++) { 
      if(dataSek[s][5]) mapSekolah[dataSek[s][5].toString().trim().toUpperCase()] = dataSek[s][6] ? dataSek[s][6].toString().trim() : ""; 
    }

    var mapDaerah = {};
    var dataDae = ambilSheetKalisRalat(ss, 'tbdaerah').getDataRange().getValues();
    for(var d = 1; d < dataDae.length; d++) { 
      if(dataDae[d][5]) mapDaerah[dataDae[d][5].toString().trim().toUpperCase()] = dataDae[d][2] ? dataDae[d][2].toString().trim() : ""; 
    }

    var mapMp = {};
    var dataMp = ambilSheetKalisRalat(ss, 'tbmp').getDataRange().getValues();
    for(var m = 1; m < dataMp.length; m++) { 
      if(dataMp[m][1]) mapMp[dataMp[m][1].toString().trim().toUpperCase()] = dataMp[m][2] ? dataMp[m][2].toString().trim() : ""; 
    }

    var dataKpmp = ambilSheetKalisRalat(ss, 'tbkpmp').getDataRange().getValues();
    if (dataKpmp.length <= 1) return { status: "FAILED", message: "Tiada rekod dijumpai dalam pangkalan data KPMP." };

    var hasilEksport = [];
    hasilEksport.push(["BIL", "KOD SEKOLAH", "NAMA SEKOLAH", "KOD PPD", "NAMA DAERAH", "KOD M/P", "NAMA MATA PELAJARAN", "NAMA KPMP", "NO KAD PENGENALAN", "JANTINA", "NO TELEFON", "EMEL DELIMA"]);

    var bil = 1;
    for (var i = 1; i < dataKpmp.length; i++) {
      var kSek = dataKpmp[i][1] ? dataKpmp[i][1].toString().trim().toUpperCase() : "";
      var kDae = dataKpmp[i][2] ? dataKpmp[i][2].toString().trim().toUpperCase() : "";
      var kMpRow = dataKpmp[i][3] ? dataKpmp[i][3].toString().trim().toUpperCase() : "";

      var padanDaerah = (kodDaerah === "SEMUA" || kDae === kodDaerah.toUpperCase());
      var padanMp = (kodMp === "SEMUA" || kMpRow === kodMp.toUpperCase());

      if (padanDaerah && padanMp) {
        var nSek = mapSekolah[kSek] || kSek;
        var nDae = mapDaerah[kDae] || kDae;
        var nMp = mapMp[kMpRow] || kMpRow;

        var namaKpmp = dataKpmp[i][4] ? dataKpmp[i][4].toString().trim() : "";
        var noKp = dataKpmp[i][5] ? dataKpmp[i][5].toString().trim() : "";
        var jantina = dataKpmp[i][6] ? dataKpmp[i][6].toString().trim() : "";
        var noTel = dataKpmp[i][7] ? dataKpmp[i][7].toString().trim() : "";
        var emel = dataKpmp[i][8] ? dataKpmp[i][8].toString().trim() : "";

        hasilEksport.push([ bil++, kSek, nSek, kDae, nDae, kMpRow, nMp, namaKpmp, noKp, jantina, noTel, emel ]);
      }
    }

    if(hasilEksport.length === 1) return { status: "FAILED", message: "Tiada rekod KPMP ditemui bagi tapisan daerah dan mata pelajaran ini." };
    return { status: "SUCCESS", data: hasilEksport };

  } catch(e) {
    return { status: "FAILED", message: e.message };
  }
}

// =========================================================================
// FUNGSI PADAM PK SELAMAT (KALIS ROW-SHIFT)
// =========================================================================
function padamPkBerdasarkanKunci(no_kp, tahun, tahap) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // Kunci trafik
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ambilSheetKalisRalat(ss, 'tbpk');
    var data = sheet.getDataRange().getValues();
    
    var rowToDelete = -1;
    // Memburu baris sebenar menggunakan 3 Kunci
    for (var i = 1; i < data.length; i++) {
      var dbTahun = data[i][1] ? data[i][1].toString().trim() : "";
      var dbTahap = data[i][4] ? data[i][4].toString().trim().toUpperCase() : "";
      var dbKp = data[i][8] ? data[i][8].toString().trim() : "";
      
      if (dbTahun === tahun.toString().trim() && 
          dbTahap === tahap.toString().trim().toUpperCase() && 
          dbKp === no_kp.toString().trim()) {
        rowToDelete = i + 1;
        break;
      }
    }
    
    if (rowToDelete > 0) {
      sheet.deleteRow(rowToDelete);
      SpreadsheetApp.flush();
      lock.releaseLock();
      return { status: "SUCCESS" };
    } else {
      lock.releaseLock();
      return { status: "FAILED", message: "Rekod tidak dijumpai (Mungkin telah dipadam oleh pengguna lain sebelum ini)." };
    }
  } catch(e) {
    if (lock) lock.releaseLock();
    return { status: "FAILED", message: e.message };
  }
}