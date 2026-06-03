function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;

    if (action === 'importData') {
      const {
        phongBan,
        projectCode,
        noiDung,
        coQuan,
        soHieu,
        ngay,
        fileName,
        mimeType,
        fileData,
        targetFolderId
      } = payload;

      let fileUrl = "";

      // 1. Lưu file vào Google Drive
      if (fileData && fileData.trim() !== '') {
        const rootFolder = DriveApp.getFolderById(targetFolderId);
        let projectFolder;

        // Tìm xem thư mục mã công trình đã tồn tại chưa
        const folders = rootFolder.getFoldersByName(projectCode);
        if (folders.hasNext()) {
          projectFolder = folders.next();
        } else {
          // Chưa tồn tại thì tạo mới
          projectFolder = rootFolder.createFolder(projectCode);
        }

        // Decode Base64 data và lưu file vào projectFolder
        const byteCharacters = Utilities.base64Decode(fileData);
        const blob = Utilities.newBlob(byteCharacters, mimeType, fileName);
        const file = projectFolder.createFile(blob);
        
        fileUrl = file.getUrl();
      }

      // 2. Ghi dữ liệu vào sheet "Data_Ho_So"
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName('Data_Ho_So');
      
      if (!sheet) {
        throw new Error('Không tìm thấy sheet Data_Ho_So');
      }

      // Ghi theo đúng cấu trúc cột
      // A: Mã CT, B: Phòng Ban, C: Nội dung văn bản, D: Số VB, E: Ngày VB, F: Link File, G: Ngày tải lên
      const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

      sheet.appendRow([
        projectCode,      // Cột A: Mã CT
        phongBan,         // Cột B: Phòng Ban
        noiDung,          // Cột C: Nội dung văn bản
        soHieu,           // Cột D: Số VB
        ngay,             // Cột E: Ngày VB
        fileUrl,          // Cột F: Link File
        timestamp         // Cột G: Ngày tải lên
      ]);

      return ContentService.createTextOutput(JSON.stringify({ 
        status: 'success', 
        message: 'Lưu dữ liệu và file thành công',
        fileUrl: fileUrl
      })).setMimeType(ContentService.MimeType.JSON);

    } else if (action === 'addCustomCategory') {
      // Logic xử lý Add Content Modal (Danh_Muc_Bo_Sung)
      const { maCongTrinh, giaiDoan, dauMuc, noiDung, phongBan } = payload;
      
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName('Danh_Muc_Bo_Sung');
      
      if (!sheet) {
        throw new Error('Không tìm thấy sheet Danh_Muc_Bo_Sung');
      }

      sheet.appendRow([
        maCongTrinh,
        giaiDoan,
        dauMuc,
        noiDung,
        phongBan
      ]);

      return ContentService.createTextOutput(JSON.stringify({ 
        status: 'success', 
        message: 'Thêm nội dung bổ sung thành công'
      })).setMimeType(ContentService.MimeType.JSON);

    } else {
      return ContentService.createTextOutput(JSON.stringify({ 
        status: 'error', 
        message: 'Không tìm thấy action phù hợp' 
      })).setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ 
      status: 'error', 
      message: error.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
