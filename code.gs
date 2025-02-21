/********************************************************
 * 1) FUNCTION TO REQUEST PERMISSIONS (Gmail, Drive, Sheets)
 ********************************************************/
function requestAllPermissions() {
  try {
    // Get the email of the active user
    var authEmail = Session.getActiveUser().getEmail();
    
    // Send a test email to trigger Gmail permissions
    MailApp.sendEmail({
      to: authEmail,
      subject: "Permissions Confirmation",
      body: "This is a test email to activate permissions for Gmail, Google Sheets, and Google Drive."
    });

    // Simple read in Sheets
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];
    var testValue = sheet.getRange("A1").getValue();

    // Access Drive
    var files = DriveApp.getFiles();

    // Display message to the user
    Browser.msgBox("Permissions requested. Please accept the permissions prompt.");
  } catch (error) {
    Logger.log("Error in requestAllPermissions: " + error.message);
  }
}

/********************************************************
 * 2) FUNCTION TO CREATE (OR REPLACE) THE INSTALLABLE TRIGGER
 ********************************************************/
function createInstallableTriggers() {
  try {
    // Delete any previous triggers using handleEdit
    var allTriggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < allTriggers.length; i++) {
      var tri = allTriggers[i];
      if (tri.getHandlerFunction() === 'handleEdit') {
        ScriptApp.deleteTrigger(tri);
      }
    }

    // Create a new installable trigger for "handleEdit"
    ScriptApp.newTrigger("handleEdit")
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onEdit()
      .create();
    
    // Display a toast message in the spreadsheet
    SpreadsheetApp.getActive().toast(
      "Installable trigger 'handleEdit' created successfully.",
      "Info",
      5
    );
  } catch (error) {
    Logger.log("Error in createInstallableTriggers: " + error.message);
  }
}

/********************************************************
 * 3) INSTALLABLE TRIGGER: FIRES ON EDITING COLUMN U
 ********************************************************/
function handleEdit(e) {
  try {
    // Ensure the edit is on the "Actualizado" sheet
    var sheet = e.range.getSheet();
    if (sheet.getName() !== "Actualizado") return;

    // Column U = 21
    var editedCell = e.range;
    var columnU = 21;

    // Only act if the edit is in column U and not on the header row
    if (editedCell.getColumn() === columnU && editedCell.getRow() > 1) {
      var selectedValue = editedCell.getValue();
      // Do nothing if the cell is empty or equals "(En blanco)"
      if (selectedValue && selectedValue !== "(En blanco)") {
        generateAndSendReport(editedCell.getRow(), selectedValue);
      }
    }
  } catch (error) {
    Logger.log("Error in handleEdit: " + error.message);
  }
}

/********************************************************
 * 4) FUNCTION TO GENERATE AND SEND THE REPORT
 ********************************************************/
function generateAndSendReport(row, selectedReport) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Actualizado");

    // Verify/create the "Logs" sheet
    var logsSheet = ss.getSheetByName("Logs");
    if (!logsSheet) {
      logsSheet = ss.insertSheet("Logs");
      // Optional headers
      logsSheet.appendRow([
        "Date & Time",
        "Name",
        "Position",
        "Email Sent",
        "Report Requested",
        "Status"
      ]);
    }

    // Basic employee data
    var site     = sheet.getRange(row, 1).getValue();  // Column A
    var name     = sheet.getRange(row, 2).getValue();  // Column B
    var position = sheet.getRange(row, 3).getValue();  // Column C
    var email    = sheet.getRange(row, 22).getValue(); // Column V = 22

    // Validate name
    if (!name) {
      Logger.log("Row " + row + " does not have an employee name.");
      return;
    }

    // Handle empty email field
    var noEmailProvided = false;
    if (!email) {
      email = "noreply@example.com";
      noEmailProvided = true;
    }

    /************************************************
     * 4.1) COMPLETE DOCUMENTS LIST (Columns)
     ************************************************/
    var allDocuments = [
      {
        name: "Certificado de Salud",         
        expiration: sheet.getRange(row, 6).getValue()   // Column F
      },
      {
        name: "Antecedentes Penales",        
        expiration: sheet.getRange(row, 8).getValue()   // Column H
      },
      {
        name: "LEY 300",                     
        expiration: sheet.getRange(row, 10).getValue()  // Column J
      },
      {
        name: "ASUME",                       
        expiration: sheet.getRange(row, 12).getValue()  // Column L
      },
      {
        name: "Departamento de la Familia",  
        expiration: sheet.getRange(row, 14).getValue()  // Column N
      },
      {
        name: "CPR",                         
        expiration: sheet.getRange(row, 16).getValue()  // Column P
      },
      {
        name: "Huellas",                     
        expiration: sheet.getRange(row, 18).getValue()  // Column R
      },
      {
        name: "Inocuidad",                   
        expiration: sheet.getRange(row, 20).getValue()  // Column T
      }
    ];

    /************************************************
     * 4.2) FILTER DOCUMENTS BASED ON SELECTION
     ************************************************/
    var selectedDocs = [];

    if (selectedReport === "Todos los Documentos") {
      selectedDocs = allDocuments;
    } else if (selectedReport === "Certificado de Salud") {
      selectedDocs = allDocuments.filter(d => d.name === "Certificado de Salud");
    } else if (selectedReport === "Antecedentes Penales") {
      selectedDocs = allDocuments.filter(d => d.name === "Antecedentes Penales");
    } else if (selectedReport === "LEY 300") {
      selectedDocs = allDocuments.filter(d => d.name === "LEY 300");
    } else if (selectedReport === "ASUME") {
      selectedDocs = allDocuments.filter(d => d.name === "ASUME");
    } else if (selectedReport === "Departamento de la Familia") {
      selectedDocs = allDocuments.filter(d => d.name === "Departamento de la Familia");
    } else if (selectedReport === "CPR") {
      selectedDocs = allDocuments.filter(d => d.name === "CPR");
    } else if (selectedReport === "Huellas") {
      selectedDocs = allDocuments.filter(d => d.name === "Huellas");
    } else if (selectedReport === "Inocuidad") {
      selectedDocs = allDocuments.filter(d => d.name === "Inocuidad");
    } else {
      // If an unexpected value arrives, do not process
      Logger.log("Unrecognized report value: " + selectedReport);
      return;
    }

    // If no documents were found, exit
    if (selectedDocs.length === 0) {
      Logger.log("No documents found for report '" + selectedReport + "'.");
      return;
    }

    /************************************************
     * 4.3) VERIFY OR CREATE A FOLDER IN DRIVE
     ************************************************/
    var folderName = "Reportes Documentos";  // Customize the folder name if desired
    var folder;
    var folders = DriveApp.getFoldersByName(folderName);
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
    }

    /************************************************
     * 4.4) CREATE A GOOGLE DOC WITH THE INFORMATION
     ************************************************/
    var docTitle = "Estado de Documentos - " + name + " (" + selectedReport + ")";
    var doc = DocumentApp.create(docTitle);
    var docId = doc.getId();
    var docBody = doc.getBody();

    // Report date
    var now = new Date();
    var formattedDate = Utilities.formatDate(now, "America/Puerto_Rico", "dd/MM/yyyy");

    // Add headings to the Doc
    docBody.appendParagraph("COMPANY")
      .setHeading(DocumentApp.ParagraphHeading.NORMAL);
    docBody.appendParagraph("Reporte de Estado de Documentos")
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);

    docBody.appendParagraph("Sitio: " + site);
    docBody.appendParagraph("Nombre: " + name);
    docBody.appendParagraph("Posición: " + position);
    docBody.appendParagraph("Reporte Solicitado: " + selectedReport);
    docBody.appendParagraph("Fecha de Reporte: " + formattedDate);

    docBody.appendHorizontalRule();
    docBody.appendParagraph("Documentos:")
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);

    // Build a table with the filtered documents
    var docTable = [["Documento", "Estado"]];
    selectedDocs.forEach(function(d) {
      docTable.push([d.name, d.expiration]);
    });
    var table = docBody.appendTable(docTable);

    // Set header row background color
    table.getRow(0).getCell(0).setBackgroundColor("#cccccc");
    table.getRow(0).getCell(1).setBackgroundColor("#cccccc");

    docBody.appendParagraph("Por favor, revisa los documentos próximos a vencer.");
    doc.saveAndClose();

    // Move the newly created Doc to the specific folder
    var file = DriveApp.getFileById(docId);
    file.moveTo(folder);

    /************************************************
     * 4.5) CONSTRUCT THE HTML EMAIL BODY WITH STYLE
     ************************************************/
    var docUrl = doc.getUrl();
    var subject = "Reporte de Documentos - " + name;

    // Additional note if no email was provided
    var extraNote = noEmailProvided 
      ? "<br><strong>Nota:</strong> No se proveyó un correo en la hoja. Se envía este reporte a " + email + " por defecto."
      : "";

    // Build the HTML table rows with alternate styling
    var tableRows = "";
    selectedDocs.forEach(function(d, index) {
      var color = (typeof d.expiration === "string" && d.expiration.indexOf("Expiró") !== -1)
                  ? "#e74c3c"  // Red
                  : "#2c3e50"; // Dark blue
      var bgColor = (index % 2 === 0) ? "#ffffff" : "#f9f9f9";

      tableRows += `
        <tr style="background-color: ${bgColor};">
          <td style="padding: 12px 15px; border-bottom: 1px solid #ecf0f1;">${d.name}</td>
          <td style="padding: 12px 15px; border-bottom: 1px solid #ecf0f1; color: ${color};">${d.expiration}</td>
        </tr>
      `;
    });

    // Construct the final HTML with a generic company logo and a problem report link
    var htmlBody = `
      <div style="font-family: 'Helvetica Neue', Arial, sans-serif; color: #333; max-width: 600px; margin: 0 auto; padding: 20px; background-color: #f8f8f8;">
        <div style="background-color: #ffffff; padding: 30px; border-radius: 5px; box-shadow: 0 2px 5px rgba(0,0,0,0.1);">
          <div style="text-align: center; margin-bottom: 20px;">
            <img src="https://via.placeholder.com/200?text=Company+Logo" width="200" alt="Company Logo" style="display:block; margin:0 auto;" />
          </div>
          <h1 style="color: #2c3e50; font-size: 24px; margin-bottom: 20px; text-align: center;">Estado de Documentos Requeridos</h1>
          <div style="margin-bottom: 20px; background-color: #ecf0f1; padding: 15px; border-radius: 5px;">
            <p style="margin: 5px 0;"><strong>Agencia:</strong> COMPANY</p>
            <p style="margin: 5px 0;"><strong>Sitio:</strong> ${site}</p>
            <p style="margin: 5px 0;"><strong>Nombre:</strong> ${name}</p>
            <p style="margin: 5px 0;"><strong>Posición:</strong> ${position}</p>
            <p style="margin: 5px 0;"><strong>Reporte solicitado:</strong> ${selectedReport}</p>
            <p style="margin: 5px 0;"><strong>Fecha de Reporte:</strong> ${formattedDate}</p>
          </div>
          <h2 style="color: #34495e; font-size: 20px; margin-top: 30px; margin-bottom: 15px;">Documentos</h2>
          <table style="border-collapse: collapse; width:100%; text-align:left; background-color: #ffffff; border-radius: 5px; overflow: hidden;">
            <tr style="background-color:#3498db; color: #ffffff;">
              <th style="padding: 12px 15px;">Documento</th>
              <th style="padding: 12px 15px;">Estado</th>
            </tr>
            ${tableRows}
          </table>
          <p style="color: #e74c3c; font-weight: bold; margin-top: 20px;">Por favor, revisa los documentos próximos a vencer.</p>
          <div style="text-align: center; margin-top: 30px;">
            <a href="${docUrl}" style="background-color: #3498db; color: #ffffff; padding: 12px 20px; text-decoration: none; border-radius: 5px; font-weight: bold;" target="_blank">Ver Reporte Completo</a>
          </div>
          <p style="font-size: 12px; color: #7f8c8d; text-align: center; margin-top: 30px;">(Generado automáticamente)</p>
          <p style="text-align: center; margin-top: 10px;">
            <a href="https://example.com/report-problem" style="font-size: 12px; color: #3498db; text-decoration: none;">Report a Problem</a>
          </p>
          ${extraNote}
        </div>
      </div>`;

    /************************************************
     * 4.6) SEND THE EMAIL
     ************************************************/
    GmailApp.sendEmail(email, subject, "", { htmlBody: htmlBody });

    /************************************************
     * 4.7) LOG THE ACTION IN "Logs"
     ************************************************/
    logsSheet.appendRow([
      new Date(),      // Date & Time
      name,            // Name
      position,        // Position
      email,           // Email
      selectedReport,  // Report requested
      "Enviado"        // Status
    ]);

    /************************************************
     * 4.8) UPDATE THE NOTE IN COLUMN U WITH SEND HISTORY 
     *      AND RESET THE CELL VALUE TO "(En blanco)"
     ************************************************/
    var noteCell = sheet.getRange(row, 21);  
    var oldNote = noteCell.getNote();        // Get existing note
    var noteDateTime = Utilities.formatDate(now, "America/Puerto_Rico", "dd/MM/yyyy HH:mm:ss");
    var newEntry = noteDateTime + " → Enviado a " + email;
    if (noEmailProvided) {
      newEntry += " (sin email en la hoja)";
    }
    var updatedNote = oldNote 
      ? oldNote + ", " + newEntry 
      : newEntry;
    noteCell.setNote(updatedNote);
    
    // Reset the triggering cell to "(En blanco)" after sending the report
    sheet.getRange(row, 21).setValue("(En blanco)");

  } catch (error) {
    Logger.log("Error in generateAndSendReport: " + error.message);
  }
}
