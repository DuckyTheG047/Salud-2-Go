function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = JSON.parse(e.postData.contents || "{}");

    const headers = [
      "timestamp","session_id","scenario","residencia","personas","necesidades",
      "gasto_medico","wtp","plan_recomendado","fit_recomendado","plan_seleccionado",
      "frecuencia_pago","precio_final","moneda","lead_guardado","celular_cotizacion",
      "correo_cotizacion","consentimiento_cotizacion","que_harias","waitlist_registrado",
      "cta_origen","nombre_waitlist","pais_waitlist","correo_waitlist","celular_waitlist",
      "interes_principal","entrevista","consentimiento"
    ];

    if (sheet.getLastRow() === 0) sheet.appendRow(headers);

    const sessionId = String(data.session_id || "").trim();
    if (!sessionId) return jsonResponse({ok:false,error:"session_id requerido"});

    const lastRow = sheet.getLastRow();
    let existingRow = null;

    if (lastRow >= 2) {
      const ids = sheet.getRange(2,2,lastRow-1,1).getValues().flat();
      const idx = ids.findIndex(v => String(v).trim() === sessionId);
      if (idx !== -1) existingRow = idx + 2;
    }

    if (existingRow) {
      const current = sheet.getRange(existingRow,1,1,headers.length).getValues()[0];
      const updated = headers.map((h,i) => {
        if (h === "timestamp") return new Date();
        if (Object.prototype.hasOwnProperty.call(data,h) && data[h] !== undefined && data[h] !== null) {
          return Array.isArray(data[h]) ? data[h].join(" | ") : data[h];
        }
        return current[i];
      });
      sheet.getRange(existingRow,1,1,headers.length).setValues([updated]);
      return jsonResponse({ok:true,action:"updated",row:existingRow});
    }

    const row = headers.map(h => {
      if (h === "timestamp") return new Date();
      const v=data[h];
      return Array.isArray(v) ? v.join(" | ") : (v ?? "");
    });
    sheet.appendRow(row);
    return jsonResponse({ok:true,action:"created",row:sheet.getLastRow()});
  } catch(err) {
    return jsonResponse({ok:false,error:String(err)});
  }
}

function doGet() {
  return jsonResponse({ok:true,message:"Salud 2 Go endpoint activo"});
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
