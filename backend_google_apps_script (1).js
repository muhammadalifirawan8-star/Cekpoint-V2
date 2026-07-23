function doPost(e) {
    try {
        // 1. Membuka koneksi ke Spreadsheet yang sedang aktif
        var doc = SpreadsheetApp.getActiveSpreadsheet();
        
        // 2. Parsing data JSON yang dikirim dari Frontend
        var payload = JSON.parse(e.postData.contents);
        
        var timestamp = new Date(); // Waktu Server
        var petugas = payload.petugas || "Unknown";
        var shift = payload.shift || "Tanpa Shift";
        var lokasi = payload.lokasi || "Unknown Checkpoint";
        var lat = payload.lat || "";
        var lng = payload.lng || "";
        var status = "Tercatat";
        var fotoUrl = "Tanpa Foto";

        // Aturan Google Sheet: Nama sheet maks 31 karakter, dan karakter tertentu tidak diizinkan.
        var namaSheet = lokasi.substring(0, 31).replace(/[\\/?*\[\]]/g, '_');
        
        // Cek apakah sheet dengan nama titik ini sudah ada
        var sheet = doc.getSheetByName(namaSheet);
        
        if (!sheet) {
            // Jika titik belum pernah di-scan sebelumnya, BUAT sheet baru
            sheet = doc.insertSheet(namaSheet);
            
            // Buat Header untuk sheet baru
            sheet.appendRow([
                "Timestamp", 
                "Nama Petugas", 
                "Shift", 
                "Titik Cek (Barcode)", 
                "Latitude", 
                "Longitude", 
                "Foto Bukti (Link)", 
                "Status"
            ]);
            
            // Styling Header agar rapi
            var headerRange = sheet.getRange("A1:H1");
            headerRange.setFontWeight("bold")
                       .setBackground("#1d4ed8") // Warna Biru Tol
                       .setFontColor("#ffffff");
            sheet.setFrozenRows(1); // Kunci baris 1 saat di-scroll
        }

        if (payload.fotoBase64) {
            try {
                // Hilangkan header data URL (data:image/jpeg;base64,)
                var base64Data = payload.fotoBase64.split(",")[1] || payload.fotoBase64;
                
                // Decode base64 menjadi file Blob
                var namaFile = "Patroli_" + namaSheet + "_" + timestamp.getTime() + ".jpg";
                var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), "image/jpeg", namaFile);
                
                // Cari atau buat folder khusus "Foto_Patroli_Tol" di Google Drive Root
                var folderName = "Foto_Patroli_Tol";
                var folders = DriveApp.getFoldersByName(folderName);
                var folder;
                
                if (folders.hasNext()) {
                    folder = folders.next();
                } else {
                    folder = DriveApp.createFolder(folderName);
                }
                
                // Simpan file foto ke dalam folder tersebut
                var file = folder.createFile(blob);
                
                // Atur agar file bisa diakses (View) oleh siapa saja yang punya link-nya 
                // (Ini berguna jika Anda ingin mengklik link foto di spreadsheet)
                file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
                
                // Ambil link url untuk dimasukkan ke sel tabel
                fotoUrl = file.getUrl();
            } catch(err) {
                fotoUrl = "Error upload foto: " + err.message;
            }
        }

        // Masukkan data ke baris paling bawah yang kosong pada sheet yang bersangkutan
        sheet.appendRow([
            timestamp, 
            petugas, 
            shift,
            lokasi, 
            lat, 
            lng, 
            fotoUrl,
            status
        ]);

        // Berikan balasan JSON sukses ke aplikasi frontend
        return ContentService.createTextOutput(JSON.stringify({ 
            "status": "success", 
            "message": "Data di titik " + namaSheet + " berhasil disimpan"
        })).setMimeType(ContentService.MimeType.JSON);

    } catch (error) {
        // Jika ada kegagalan teknis, kembalikan respons error
        return ContentService.createTextOutput(JSON.stringify({ 
            "status": "error", 
            "message": error.toString() 
        })).setMimeType(ContentService.MimeType.JSON);
    }
}