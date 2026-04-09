/* Updated and maintained by internetgeeks.org */
function onEdit(event) { 
  var timezone = "GMT-6";
  var timestamp_format = "dd/MM/yyyy HH:mm:ss"; // Timestamp Format. 
  var updateColName = "2 do EMPACADOR";
  var timeStampColName = "FECHA FIN EMPAQUE";

  var sheet = event.source.getActiveSheet();
  var actRng = event.source.getActiveRange();
  var editColumn = actRng.getColumn();
  var index = actRng.getRowIndex();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
  var dateCol = headers[0].indexOf(timeStampColName);
  var updateCol = headers[0].indexOf(updateColName); 
  updateCol = updateCol + 1;

  if (dateCol > -1 && index > 1 && editColumn == updateCol) { 
    var cell = sheet.getRange(index, dateCol + 1);
    var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
    cell.setValue(date);
  }
}

/* ===========================================================
   CONSULTA DE OP EN PESTAÑAS BITACORA POR AÑO + ZPL
   =========================================================== */

const SPREADSHEET_ID = '1eskl20JUEFl7BDgughFxnnQJ_HI60qtSz1fSkxv36wY';

/** Utils base */
function _getSs() { 
  return SpreadsheetApp.openById(SPREADSHEET_ID); 
}

/** Normaliza texto: mayúsculas, sin acentos, sin dobles espacios */
function _norm(s) {
  return String(s || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // quita acentos
    .toUpperCase()
    .replace(/\s+/g, ' ')
    .trim();
}

/** Extrae un año 20xx de la OP (ej. OP-2025-000123 -> 2025) */
function extractYearFromOP(op) {
  const m = String(op || '').match(/\b(20\d{2})\b/);
  return m ? m[1] : null;
}

/** Lista de hojas cuyo nombre empieza por BITACORA (insensible a acentos/mayúsculas) */
function getBitacoraSheets(ss) {
  const all = ss.getSheets();
  return all.filter(s => /^BITACORA\b/.test(_norm(s.getName())));
}

/** Devuelve la hoja que mejor coincide con el año: "BITACORA {AÑO}" (insensible) */
function matchBitacoraByYear(ss, year) {
  if (!year) return null;
  const target = `BITACORA ${year}`;

  const hit = ss.getSheets().find(s => _norm(s.getName()) === _norm(target));
  if (hit) return hit;

  const contains = ss.getSheets().find(s => _norm(s.getName()).includes(_norm(target)));
  return contains || null;
}

/** Candidatas ordenadas: 1) BITACORA {AÑO}, 2) todas las BITACORA, 3) resto */
function getCandidateSheets(op) {
  const ss = _getSs();
  const year = extractYearFromOP(op);
  const preferred = matchBitacoraByYear(ss, year);
  const bitacoras = getBitacoraSheets(ss);

  const set = new Set();
  if (preferred) set.add(preferred.getName());
  bitacoras.forEach(s => set.add(s.getName()));
  ss.getSheets().forEach(s => set.add(s.getName())); // completa con el resto

  return Array.from(set).map(name => ss.getSheetByName(name));
}

/** Busca una OP exacta en col B y devuelve Código (D) y Descripción (E) */
function getOPInfo(opRaw) {
  const op = String(opRaw || '').trim();
  if (!op) return { ok: false, msg: 'Ingresa una OP.', data: null };

  const sheets = getCandidateSheets(op);
  for (const sh of sheets) {
    const lastRow = sh.getLastRow();
    if (lastRow < 2) continue;

    const rangeColB = sh.getRange(2, 2, lastRow - 1, 1); // B desde fila 2
    const cell = rangeColB.createTextFinder(op)
      .matchCase(false)
      .matchEntireCell(true)   // exacto en la celda
      .findNext();

    if (!cell) continue;

    const row = cell.getRow();
    const codigo = sh.getRange(row, 4).getDisplayValue();      // D
    const descripcion = sh.getRange(row, 5).getDisplayValue(); // E

    return {
      ok: true,
      msg: 'OK',
      data: { 
        op,
        codigo,
        descripcion,
        fila: row,
        hoja: sh.getName()
      }
    };
  }

  return { ok: false, msg: `OP "${op}" no encontrada en ninguna pestaña.`, data: null };
}

/** Escapa caracteres conflictivos para ZPL (^, ~, \) */
function zplSanitize(text) {
  return String(text || '')
    .replace(/[\^~\\]/g, ' ')
    .replace(/\r?\n/g, ' ')
    .trim();
}

/* ===========================================================
   ZPL: ETIQUETA 1 (ROLLO 40x22)
   =========================================================== */
function generateZPLLabel1(opRaw, batchRaw, shiftRaw, qtyRaw) {
  const info = getOPInfo(opRaw);
  if (!info || !info.ok) {
    return { ok: false, msg: info && info.msg ? info.msg : 'No se encontró la OP.', data: null };
  }

  const d = info.data || {};
  const codigo      = zplSanitize(d.codigo);
  const descripcion = zplSanitize(d.descripcion);

  const batch = zplSanitize(batchRaw);
  const shift = zplSanitize(shiftRaw);
  const qty   = zplSanitize(qtyRaw);

  if (!batch) return { ok: false, msg: 'Batch vacío. Captúralo en la app.', data: null };
  if (!shift) return { ok: false, msg: 'Shift vacío. Captúralo en la app.', data: null };
  if (!qty)   return { ok: false, msg: 'QTY vacío. Captúralo en la app.', data: null };

  // QR SOLO con el código de producto
  const qrData = codigo;

  const zplLines = [
    '^XA',
    '^PW320',
    '^LL185',
    '^LH0,0',
    '^CI28',

    '^FO0,10^A0N,50,20^FDPart No.^FS',
    '^FO0,10^A0N,60,20^FD' + codigo + '^FS',

    '^FO0,40^A0N,50,18^FD' + descripcion + '^FS',

    '^FO0,125^A0N,50,18^FDQty: ' + qty + ' Labels^FS',
    '^FO0,135^A0N,50,18^FDBatch: ' + batch + '^FS',
    '^FO0,145^A0N,50,18^FDShift: ' + shift + '^FS',

    '^FO165,80^BQN,3,4',
    '^FDLA,' + qrData + '^FS',

    '^XZ'
  ];

  return {
    ok: true,
    msg: 'ZPL generado.',
    data: {
      op: d.op,
      codigo: d.codigo,
      descripcion: d.descripcion,
      qty, batch, shift,
      hoja: d.hoja,
      fila: d.fila,
      zpl: zplLines.join('\n')
    }
  };
}

/* ===========================================================
   ZPL: ETIQUETA PISA 102x152
   (usa tu layout ZPL y reemplaza variables)
   =========================================================== */
function generateZPLLabelPisa(opRaw, fechaRaw, epRaw, idPisaRaw, codBarrasRaw, descClienteRaw, loteRaw, rollosCajaRaw, millaresCajaRaw, qrRaw) {
  const info = getOPInfo(opRaw);
  if (!info || !info.ok) {
    return { ok: false, msg: info && info.msg ? info.msg : 'No se encontró la OP.', data: null };
  }

  const base = info.data || {};

  const FECHA        = zplSanitize(fechaRaw);
  const EP           = zplSanitize(epRaw);
  const ID_PISA      = zplSanitize(idPisaRaw);
  const COD_BARRAS   = zplSanitize(codBarrasRaw);
  const DESC_CLIENTE = zplSanitize(descClienteRaw);
  const LOTE         = zplSanitize(loteRaw);

  const CODIGO       = zplSanitize(base.codigo);
  const DESCRIPCION  = zplSanitize(base.descripcion);

  const ROLLOS_CAJA   = zplSanitize(rollosCajaRaw);
  const MILLARES_CAJA = zplSanitize(millaresCajaRaw);

  // QR: si no te mandan nada, usa el código
  const QR = zplSanitize(qrRaw || base.codigo);

  if (!FECHA)        return { ok:false, msg:'Fecha vacía.', data:null };
  if (!EP)           return { ok:false, msg:'EP vacío.', data:null };
  if (!ID_PISA)      return { ok:false, msg:'ID PISA vacío.', data:null };
  if (!COD_BARRAS)   return { ok:false, msg:'Código de barras vacío.', data:null };
  if (!DESC_CLIENTE) return { ok:false, msg:'Desc. Cliente vacía.', data:null };
  if (!LOTE)         return { ok:false, msg:'Lote vacío.', data:null };
  if (!ROLLOS_CAJA)  return { ok:false, msg:'Rollos por caja vacío.', data:null };
  if (!MILLARES_CAJA)return { ok:false, msg:'Millares por caja vacío.', data:null };

  // 🔥 Tu ZPL tal cual (con placeholders)
  const zplTemplate = `
^XA

^FO100,20^GFA,4686,4686,33,,::::::::::::::::::N07FF8,M07JFC,L03LF,L0MFE,K03NF,K07NFC,J01OFE,J03PF8,J07PFC,J0QFE,I01RF,I03RF8,I07RFC,I0SFE,001SFE,001LF03LF,003KFI03KF8,007JFCJ0KF8,007JFK03JFC,00JFEK01JFC,00JF8L0JFE,00JFM07IFE,00IFEM03JF,00IFEM01JF,00IFCN0JF,00IF8I0FCI0JF8,00IF8007FFI07IF8,00IF800IFC007IF8,00IF001IFE003IF8,00IF003JF003IFC,00IF007JF003IFC,00FFE007JF801IFC,00FFE00KF801IFCK01FE07F81FF87F807F81FF807801FE381FF03C07,00FFE00KFC01IFCK07FE1FFC1FFC7FE1FFC1FFC07C03FE383FF83E07,00FFE00KFC01IFCK0F803C1E1C1C70E3E1E1C1C0FC078038783C3E07,00FFE00KFC01IFCJ01E00780F1C1C70E780F1C1E1DC0F0038F01E3F07,00FFE00KFC01IFCJ01E0070071C1C60E70071C1C1CE1E0038E00E3F86,00FFE00KFC01IFCJ01C0070079C1C61E70079C1C18E1C0079E00E3B86,00FFE00KFC01IFCJ01C0070079DF86FC70079DF83861C0079C00E39C6,00FFE00KFC01IFCJ01C0070079DE06F870079DE07FF1C0079C00E39E6,00IF00KF801IFCJ01C0070079CE060070079CE07FF1C0079E00E38E6,00IF007JF801IFCJ01C0070071CE060070071C70F079E0079E00E30F6,00IF007JF001IFCJ01E00780F1C70E00780F1C70E039E0078E01E307E,00IF803JF001IFCK0F003C1E1C78E003C1E1C79E038F0078F83C303E,00IF801IFE001IFCK07FE1FFC1C38E001FFE1C39C03C7FE707FF8703E,00IFC00IFC001IFCK03FC0FF81C3CEI0FF81C3FC01C3FE701FF0701E,00IFC003FFJ07FF8,00IFEI0F8,00JF,00JF8,00JFC,007JF,007JF8,003KF,001KFE,001gJF8K07FF9FF80FEIF3FE01FC0780E380FF,I0gJF8K07FF9FF83FEIF3FF07FF07C0E383FF,I07gIF8L0781C007C00F03878F8787E0E387C,I03gIF8L0781C00FI0703839E03C7E0E3878,I01gIF8L0781C01EI0703839C01C7F0E38F,J0gIF8L0781C01EI0E03871C01C770E38E,J07gHF8L0781FF1CI0E039F3C01E738E38E,J03gHF8L0781FF1CI0E07BC3C01E73CE38E,J01gHF8L0701C01CI0E07B83C01C71CE38E,K07gGF8L0703C01EI0E071C1C01C70EE38E,K03gGF8L0703800EI0E071C1C03C70FE38F,L0gGF8L0703800FI0E070E1F078707C3878,L03gF8L0703FF87FE0E070E0IF0707C383FE,M07YF8L0703FF83FE0E070703FE0703C381FE,N07PFgH078N07O03C,U0JF,::::::U0JF007KF8V0E006I03F0FC0078102,U0JF007KF8V0E006I03F0FC0078102,^FS

^PW1212
^LL812
^LH0,0
^CF0,30

^FO20,170^GB1170,620,3^FS

^FO700,40^A0N,40,40^FDFecha de Entrada^FS
^FO760,80^A0N,40,40^FD$FECHA^FS

^FO700,120^A0N,30,30^FDwww.tectronic.com.mx^FS

^FO610,170^GB580,130,3^FS

^FO40,210^A0N,45,45^FDID PISA:^FS
^FO220,210^A0N,45,45^FD$ID_PISA^FS

^FO750,190^BCN,70,Y,N,N
^FD$COD_BARRAS^FS

^FO20,300^GB1170,1,3^FS

^FO40,330^A0N,32,32^FDDesc. Cliente:^FS
^FO250,330^A0N,32,32^FD$DESC_CLIENTE^FS

^FO20,420^GB600,100,2^FS
^FO620,420^GB570,100,2^FS

^FO40,440^A0N,30,30^FDCodigo:^FS
^FO40,480^A0N,30,30^FD$CODIGO^FS

^FO650,450^A0N,55,55^FDLote:^FS
^FO780,450^A0N,55,55^FD$LOTE^FS

^FO20,520^GB1170,1,3^FS

^FO40,550^A0N,32,32^FDDescripcion:^FS
^FO250,550^A0N,32,32^FD$DESCRIPCION^FS

^FO20,640^GB400,150,2^FS

^FO40,650^A0N,30,30^FDRollos por caja^FS
^FO40,690^A0N,30,30^FD$ROLLOS_CAJA^FS

^FO420,640^GB400,150,2^FS

^FO500,650^A0N,30,30^FDMillares por caja^FS
^FO500,690^A0N,40,40^FD$MILLARES_CAJA^FS

^FO820,640^GB370,150,2^FS

^FO950,640^BQN,2,6
^FDQA,$QR^FS

^XZ
`.trim();

  // Reemplaza placeholders
  const zpl = zplTemplate
    .replaceAll('$FECHA', FECHA)
    .replaceAll('$ID_PISA', ID_PISA)
    .replaceAll('$COD_BARRAS', COD_BARRAS)
    .replaceAll('$DESC_CLIENTE', DESC_CLIENTE)
    .replaceAll('$CODIGO', CODIGO)
    .replaceAll('$LOTE', LOTE)
    .replaceAll('$DESCRIPCION', DESCRIPCION)
    .replaceAll('$ROLLOS_CAJA', ROLLOS_CAJA)
    .replaceAll('$MILLARES_CAJA', MILLARES_CAJA)
    .replaceAll('$QR', QR);

  return {
    ok: true,
    msg: 'ZPL PISA generado.',
    data: {
      op: base.op,
      codigo: base.codigo,
      descripcion: base.descripcion,
      fecha: FECHA,
      ep: EP,
      idPisa: ID_PISA,
      codBarras: COD_BARRAS,
      descCliente: DESC_CLIENTE,
      lote: LOTE,
      rollosCaja: ROLLOS_CAJA,
      millaresCaja: MILLARES_CAJA,
      qr: QR,
      hoja: base.hoja,
      fila: base.fila,
      zpl
    }
  };
}

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Consulta OP')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
