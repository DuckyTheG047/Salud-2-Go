function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = JSON.parse(e.postData.contents || "{}");

    const headers = [
      "timestamp","session_id","scenario","residencia","personas","necesidades",
      "gasto_medico","wtp","plan_recomendado","fit_recomendado","plan_seleccionado",
      "frecuencia_pago","precio_final","moneda","lead_guardado","tipo_contacto",
      "contacto","que_harias","waitlist_registrado","cta_origen","nombre_waitlist",
      "pais_waitlist","correo_waitlist","celular_waitlist","interes_principal",
      "entrevista","consentimiento"
    ];

    if (sheet.getLastRow() === 0) sheet.appendRow(headers);

    const sessionId = String(data.session_id || "").trim();
    if (!sessionId) return jsonResponse({ok:false,error:"session_id requerido"});

    const lastRow = sheet.getLastRow();
    let existingRow = null;

    if (lastRow >= 2) {
      const sessionValues = sheet.getRange(2, 2, lastRow - 1, 1).getValues().flat();
      const index = sessionValues.findIndex(v => String(v).trim() === sessionId);
      if (index !== -1) existingRow = index + 2;
    }

    if (existingRow) {
      const currentValues = sheet.getRange(existingRow, 1, 1, headers.length).getValues()[0];
      const updatedValues = headers.map((header, i) => {
        if (header === "timestamp") return new Date();
        if (Object.prototype.hasOwnProperty.call(data, header) &&
            data[header] !== undefined && data[header] !== null) {
          return Array.isArray(data[header]) ? data[header].join(" | ") : data[header];
        }
        return currentValues[i];
      });
      sheet.getRange(existingRow, 1, 1, headers.length).setValues([updatedValues]);
      return jsonResponse({ok:true,action:"updated",row:existingRow});
    }

    const newRow = headers.map(header => {
      if (header === "timestamp") return new Date();
      const value = data[header];
      return Array.isArray(value) ? value.join(" | ") : (value !== undefined && value !== null ? value : "");
    });

    sheet.appendRow(newRow);
    return jsonResponse({ok:true,action:"created",row:sheet.getLastRow()});

  } catch (error) {
    return jsonResponse({ok:false,error:String(error)});
  }
}

function doGet() {
  return jsonResponse({ok:true,message:"Salud 2 Go endpoint activo"});
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
