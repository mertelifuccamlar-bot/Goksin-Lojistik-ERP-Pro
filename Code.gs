function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('GÖKŞİN-GÖKŞEN LOJİSTİK SİSTEMİ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// GÜVENLİ GİRİŞ KONTROLÜ
function loginKontrol(kullaniciAdi, sifre) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Kullanicilar") || ss.getSheets()[0];
    var data = sheet.getDataRange().getValues();
    kullaniciAdi = kullaniciAdi.toUpperCase().trim();
    for(var i = 1; i < data.length; i++) {
      if(data[i][0] && data[i][0].toString().toUpperCase() == kullaniciAdi && data[i][1].toString() == sifre) {
        var token = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, kullaniciAdi + Date.now()));
        return { durum: "BAŞARILI", isim: data[i][2] ? data[i][2].toString().toUpperCase() : kullaniciAdi, yetkiler: data[i][3] ? data[i][3].toString() : "KONTROL_PANELI", token: token };
      }
    }
  } catch(e) { return {durum: "HATA", mesaj: e.toString()}; }
  return {durum: "HATALI"};
}

function sonrakiIDOlustur(sayfaAdi, önEk) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sayfaAdi);
    if (!sheet) return önEk + "-1001";
    var sonSatir = sheet.getLastRow();
    if (sonSatir < 2) return önEk + "-1001";
    var sonID = sheet.getRange(sonSatir, 1).getValue().toString();
    var parcalar = sonID.split("-");
    if (parcalar.length < 2) return önEk + "-1001";
    return önEk + "-" + (parseInt(parcalar[1]) + 1);
  } catch(e) { return önEk + "-1001"; }
}

function dinamikVeriEkle(sayfaAdi, önEk, hamVeriDizisi) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sayfaAdi);
    if (!sheet) return "SAYFA BULUNAMADI";
    var yeniID = sonrakiIDOlustur(sayfaAdi, önEk);
    hamVeriDizisi[0] = yeniID;
    var temizVeri = hamVeriDizisi.map(function(item) {
      if (item instanceof Date) return Utilities.formatDate(item, Session.getScriptTimeZone(), "yyyy-MM-dd");
      return typeof item === 'string' ? item.toUpperCase().trim() : item;
    });
    sheet.appendRow(temizVeri);
    return "KAYIT BAŞARILI: " + yeniID;
  } catch(e) { return "HATA: " + e.toString(); }
}

function getTabloVerileri(sayfaAdi) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sayfaAdi);
    if(!sheet) return [];
    var range = sheet.getDataRange();
    if (range.getNumRows() < 2) return [];
    var data = range.getValues();
    for(var i=1; i<data.length; i++) {
      for(var j=0; j<data[i].length; j++) {
        if(data[i][j] instanceof Date) data[i][j] = Utilities.formatDate(data[i][j], Session.getScriptTimeZone(), "yyyy-MM-dd");
        else data[i][j] = data[i][j].toString();
      }
    }
    return data;
  } catch(e) { return []; }
}

function getTümSistemVerileri() {
  return {
    rotalar: getTabloVerileri("Rotalar"),
    iadeler: getTabloVerileri("Iadeler"),
    izinler: getTabloVerileri("Izinler"),
    filo: getTabloVerileri("Filo")
  };
}

function alimSatirGuncelle(id, islemTuru, bayiAdi, tarih, miktar, sofor) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Iadeler");
    var data = sheet.getDataRange().getValues();
    for(var i=1; i<data.length; i++) {
      if(data[i][0].toString() == id.toString()) {
        sheet.getRange(i+1, 2).setValue(islemTuru.toUpperCase().trim());
        sheet.getRange(i+1, 3).setValue(bayiAdi.toUpperCase().trim());
        sheet.getRange(i+1, 4).setValue(tarih);
        sheet.getRange(i+1, 5).setValue(miktar);
        sheet.getRange(i+1, 6).setValue(sofor.toUpperCase().trim());
        return "ALIM GÜNCELLEME BAŞARILI";
      }
    }
  } catch(e) { return "HATA"; }
  return "KAYIT BULUNAMADI";
}

function izinSatirGuncelle(id, personel, izinTuru, tarih, saat, aciklama) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Izinler");
    var data = sheet.getDataRange().getValues();
    for(var i=1; i<data.length; i++) {
      if(data[i][0].toString() == id.toString()) {
        sheet.getRange(i+1, 2).setValue(personel.toUpperCase().trim());
        sheet.getRange(i+1, 3).setValue(izinTuru.toUpperCase().trim());
        sheet.getRange(i+1, 4).setValue(tarih);
        sheet.getRange(i+1, 5).setValue(saat ? saat.trim() : "-");
        sheet.getRange(i+1, 6).setValue(aciklama.toUpperCase().trim());
        return "İZİN GÜNCELLEME BAŞARILI";
      }
    }
  } catch(e) { return "HATA"; }
  return "KAYIT BULUNAMADI";
}

function gorevTamamla(sayfaAdi, id) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sayfaAdi);
    var data = sheet.getDataRange().getValues();
    var sutunIndex = (sayfaAdi === "Filo") ? 6 : 7;
    for(var i=1; i<data.length; i++) {
      if(data[i][0].toString() == id.toString()) { sheet.getRange(i+1, sutunIndex).setValue("TAMAMLANDI"); break; }
    }
  } catch(e) {}
  return "İŞLEM TAMAMLANDI";
}

function kesinSil(sayfaAdi, id) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sayfaAdi);
    var data = sheet.getDataRange().getValues();
    for(var i=data.length-1; i>=1; i--) {
      if(data[i][0].toString() == id.toString()) { sheet.deleteRow(i+1); break; }
    }
  } catch(e) {}
  return "KAYIT SİLİNDİ";
}

function durumGuncelle(sayfaAdi, id, yeniDurum, ertelemeTarihi) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sayfaAdi);
    var data = sheet.getDataRange().getValues();
    for(var i=1; i<data.length; i++) {
      if(data[i][0].toString() == id.toString()) {
        sheet.getRange(i+1, 7).setValue(yeniDurum.toUpperCase());
        if(ertelemeTarihi) sheet.getRange(i+1, 8).setValue(ertelemeTarihi);
        break;
      }
    }
  } catch(e) {}
  return "GÜNCELLENDİ";
}

function soforGuncelle(sayfaAdi, id, yeniSofor) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sayfaAdi);
    var data = sheet.getDataRange().getValues();
    for(var i=1; i<data.length; i++) {
      if(data[i][0].toString() == id.toString()) { sheet.getRange(i+1, 6).setValue(yeniSofor.toUpperCase()); break; }
    }
  } catch(e) {}
  return "ŞOFÖR GÜNCELLENDİ";
}

function sifreDegistir(kullaniciIsmi, yeniSifre) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Kullanicilar");
    var data = sheet.getDataRange().getValues();
    for(var i = 1; i < data.length; i++) {
      if(data[i][2].toUpperCase() == kullaniciIsmi.toUpperCase()) { sheet.getRange(i+1, 2).setValue(yeniSifre); return "BAŞARILI"; }
    }
  } catch(e) {}
  return "BULUNAMADI";
}

// 🛡️ DİNAMİK PERSONEL & ŞOFÖR YÖNETİM MERKEZİ (VERİTABANI ODAKLI)
function getSistemVerileri() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pSheet = ss.getSheetByName("Personeller");
  var fSheet = ss.getSheetByName("Filo_Araclar");
  var bSheet = ss.getSheetByName("Bolgeler");
  
  // Eğer sekmeler yoksa otomatik oluştur
  if(!pSheet) { pSheet = ss.insertSheet("Personeller"); pSheet.appendRow(["ID", "TUR", "ISIM"]); }
  if(!fSheet) { fSheet = ss.insertSheet("Filo_Araclar"); fSheet.appendRow(["ID", "PLAKA"]); }
  if(!bSheet) { bSheet = ss.insertSheet("Bolgeler"); bSheet.appendRow(["ID", "BOLGE"]); bSheet.appendRow(["RG-1001", "BURSA MERKEZ"]); bSheet.appendRow(["RG-1002", "MARMARA BÖLGESİ"]); }

  // İlk Kurulumda Sabit Verilerin Mükerrersiz Basılması (Tablo boşsa çalışır)
  if(pSheet.getLastRow() < 2) {
    var sabitKadro = [
      ["PERSONEL", "İBRAHİM YAYLA"], ["PERSONEL", "SAVAŞ KURT"], ["PERSONEL", "CÜNEYT YAMAN"], 
      ["PERSONEL", "ATİLLA GÜNDÜZ"], ["PERSONEL", "AHMET KALLEM"], ["PERSONEL", "İSMAİL AKTAŞ"], 
      ["PERSONEL", "HASAN KURT"], ["PERSONEL", "EMİR YILDIZLI"], ["PERSONEL", "ERMAN ÇALIBAŞI"], 
      ["PERSONEL", "NAZIM BAKAN"], ["PERSONEL", "İBRAHİM ŞAHİN"],
      ["SOFOR", "MERT ÜÇÇAMLAR"], ["SOFOR", "AHMET AKPINAR"], ["SOFOR", "AHMET YANLIZ"], 
      ["SOFOR", "HÜSEYİN YURTOĞLU"], ["SOFOR", "SAVAŞ SEVİM"], ["SOFOR", "MUSTAFA MERT SAVAŞ"], 
      ["SOFOR", "EMİRHAN ARSLAN"], ["SOFOR", "GÖKHAN AKDOĞAN"], ["SOFOR", "AHMET HAN"], 
      ["SOFOR", "BORA DOĞAN"], ["SOFOR", "MUSA KARABULUT"], ["SOFOR", "SERHAT İNAK"], 
      ["SOFOR", "ENDER AKBULUT"], ["SOFOR", "EROL ŞENTÜRK"]
    ];
    sabitKadro.forEach(function(item, idx) {
      pSheet.appendRow(["PR-" + (1001 + idx), item[0], item[1].toUpperCase()]);
    });
  }

  if(fSheet.getLastRow() < 2) {
    var sabitPlakalar = [
      "16BZN015", "16BMY363", "16BMY409", "16BMY365", "16BPF265", 
      "16BZD941", "16BKD329", "16BKD330", "16BPF271", "16ANP979", 
      "16BMY417", "16AKR991", "16BMY408", "16BZD949", "16ADL111"
    ];
    sabitPlakalar.forEach(function(plk, idx) {
      fSheet.appendRow(["CAR-" + (1001 + idx), plk.toUpperCase()]);
    });
  }

  // 🔄 Veritabanından (Excel) Canlı Okuma Motoru ve Mükerrer Temizliği
  var personeller = [], soforler = [], plakalar = [], bolgeler = [];
  
  var pData = pSheet.getDataRange().getValues();
  var kontrolSeti = {}; // Mükerrer engelleme hafızası
  
  for(var i = 1; i < pData.length; i++) {
    if(!pData[i][1] || !pData[i][2]) continue;
    var tur = pData[i][1].toString().toUpperCase().trim();
    var isim = pData[i][2].toString().toUpperCase().trim();
    
    var anahtar = tur + "_" + isim;
    if(!kontrolSeti[anahtar]) {
      kontrolSeti[anahtar] = true;
      if(tur === "PERSONEL" || tur === "İÇ PERSONEL") {
        if(personeller.indexOf(isim) === -1) personeller.push(isim);
      }
      if(tur === "SOFOR" || tur === "ŞOFÖR") {
        if(soforler.indexOf(isim) === -1) soforler.push(isim);
        // Şoförler aynı zamanda personel listesinde de yer almalı
        if(personeller.indexOf(isim) === -1) personeller.push(isim);
      }
    }
  }

  var fData = fSheet.getDataRange().getValues();
  for(var j = 1; j < fData.length; j++) {
    if(fData[j][1]) {
      var p = fData[j][1].toString().toUpperCase().trim();
      if(plakalar.indexOf(p) === -1) plakalar.push(p);
    }
  }

  var bData = bSheet.getDataRange().getValues();
  for(var k = 1; k < bData.length; k++) {
    if(bData[k][1]) {
      var b = bData[k][1].toString().toUpperCase().trim();
      if(bolgeler.indexOf(b) === -1) bolgeler.push(b);
    }
  }

  return {
    personeller: personeller.sort(),
    soforler: soforler.sort(),
    plakalar: plakalar.sort(),
    bolgeler: bolgeler.sort()
  };
}
/****************************************************
 * PERSONEL YÖNETİM MERKEZİ
 ****************************************************/

function personelEkle(tur, isim) {

  tur = String(tur).toUpperCase().trim();
  isim = String(isim).toUpperCase().trim();

  if (!isim) return "İSİM BOŞ OLAMAZ";

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sh = ss.getSheetByName("Personeller");

  if (!sh) {
    sh = ss.insertSheet("Personeller");
    sh.appendRow(["ID","TUR","ISIM"]);
  }

  const data = sh.getDataRange().getValues();

  for (let i=1;i<data.length;i++){

    if(
      data[i][1].toString().toUpperCase()==tur &&
      data[i][2].toString().toUpperCase()==isim
    ){
      return "BU KAYIT ZATEN VAR";
    }

  }

  const id="PR-"+Utilities.getUuid().substring(0,8).toUpperCase();

  sh.appendRow([id,tur,isim]);

  return "KAYIT BAŞARILI";

}



/****************************************************
 PERSONEL SİL
 ****************************************************/

function personelSil(isim){

  isim=String(isim).toUpperCase().trim();

  const sh=SpreadsheetApp
  .getActiveSpreadsheet()
  .getSheetByName("Personeller");

  const data=sh.getDataRange().getValues();

  for(let i=data.length-1;i>=1;i--){

    if(data[i][2].toString().toUpperCase()==isim){

      sh.deleteRow(i+1);

      return "SİLİNDİ";

    }

  }

  return "BULUNAMADI";

}
