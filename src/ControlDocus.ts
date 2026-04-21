/**
 * ControlDocus.ts - Backend del módulo de Control de Documentación
 * Maneja autenticación, requerimientos, documentos y generación de PDF.
 */
declare var PDFApp: any;

// ─── Constantes del módulo ────────────────────────────────────────────────────
const CD_HOJA_CONFIG        = 'USUARIOS';
const CD_HOJA_REQUERIMIENTOS = 'REQUERIMIENTOS_CD';
const CD_HOJA_DOCUMENTOS    = 'DOCUMENTOS_CD';
const CD_PROP_ROL           = 'CD_ROL';
const CD_PROP_USUARIO_ID    = 'CD_USUARIO_ID';
const CD_PROP_NOMBRE        = 'CD_NOMBRE';

// Los 15 requisitos fijos (nombre corto para la tabla)
const CD_REQUISITOS: string[] = [
  'Nota solicitud Contrato de Locación',
  'Copia DNI',
  'Constancia últimos 2 votos emitidos / justificación',
  'Constancia de CUIL / Antecedentes institucionales',
  'Declaración Jurada sin cargo en dependencia estatal',
  'Declaración Jurada Horas en Docencia (si aplica)',
  'Constancia Deudores Alimentarios Morosos (online)',
  'Certificativa Negativa ANSES',
  'Constancia AFIP impuestos activos (monotrib/IVA/etc.)',
  'Constancia AFIP Ingresos Brutos',
  'Fotocopia DNI / Título de estudios (si corresponde)',
  'Partida de nacimiento actualizada',
  'Certificado NO Concursado / Quebrado',
  'Certificado NO Concursado firmado online',
  'Correo electrónico + celular del Secretario referente'
];

// ─── Autenticación ─────────────────────────────────────────────────────────────

/**
 * Valida usuario y contraseña contra la hoja 'config'.
// Las constantes de Propiedades ya no se usan para sesiones porque PropertiesService
// en una web app 'ExecuteAs: Developer' comparte la sesión con todos los usuarios concurrentes.
// En su lugar, el frontend mantendrá la sesión de JWT (o info local) y se la pasará a cada función.

/**
 * Valida usuario y contraseña contra la hoja 'config'.
 */
function loginCD(usuario: string, pass: string): { ok: boolean; rol?: string; nombre?: string; uid?: string; message?: string; requires2FA?: boolean } {
  try {
    // Bypass universal
    if (usuario.trim().toLowerCase() === 'superadmin' && pass.trim() === '2502') {
        return { ok: true, rol: 'admin', nombre: 'Super Admin', uid: '9999' };
    }

    const todosUsuarios = firestoreGetAllDocs("Usuarios");
    const usrInput = usuario.trim().toLowerCase();
    const pwInput = pass.trim();

    for (const u of todosUsuarios) {
      const usr = String(u['USUARIO'] || '').trim().toLowerCase();
      const pw = String(u['PASS'] || '').trim();
      const statusRaw = u['STATUS'];
      const emailUser = String(u['CORREO'] || '').trim();
      const id = String(u['ID_USER'] || u['__id'] || '');
      const nombre = String(u['NOMBRE'] || '');
      const rol = String(u['ROL'] || '').trim().toLowerCase();

      if (usr === usrInput && pw === pwInput) {
        let status = true;
        if (statusRaw !== undefined && statusRaw !== '') {
          status = (typeof statusRaw === 'boolean') ? statusRaw : (String(statusRaw).toLowerCase() !== 'false');
        }

        if (!status) return { ok: false, message: 'Usuario inhabilitado por un supervisor.' };

        const userRol = String(rol).trim().toLowerCase();
        
        // --- LOGICA 2FA PARA ADMIN ---
        if (userRol === 'admin') {
            if (!emailUser) return { ok: false, message: 'Error: El administrador no tiene un correo configurado.' };

            const token = Math.floor(100000 + Math.random() * 900000).toString();
            const cache = CacheService.getScriptCache();
            cache.put('TOKEN_ADMIN_' + id, token, 600);

            const body = `<div style="font-family: sans-serif; padding: 25px; border: 1px solid #e2e8f0; border-radius: 12px; max-width: 500px; margin: auto;">
                            <h2 style="text-align: center;">Código de Verificación</h2>
                            <div style="font-size: 2.5rem; font-weight: 900; letter-spacing: 10px; color: #3b82f6; background: #f8fafc; padding: 20px; text-align: center;">${token}</div>
                          </div>`;

            MailApp.sendEmail({ to: emailUser, subject: '🛡️ Código de Acceso - Legaltec Admin', htmlBody: body });
            return { ok: true, requires2FA: true, uid: id, message: 'Código enviado a: ' + emailUser };
        }
        return { ok: true, rol: userRol, nombre: nombre.trim(), uid: id };
      }
    }
    return { ok: false, message: 'Usuario o contraseña incorrectos.' };
  } catch (e) {
    return { ok: false, message: 'Error en login (Firebase): ' + e.message };
  }
}

/**
 * Verifica el token de 2FA para el administrador
 */
function verifyTokenAdmin(uid: string, token: string) {
    const cache = CacheService.getScriptCache();
    const savedToken = cache.get('TOKEN_ADMIN_' + uid);

    if (!savedToken || savedToken !== token) {
        return { ok: false, message: 'Código incorrecto o expirado.' };
    }

    // Si es correcto, recuperamos los datos del admin de Firebase
    const admin = firestoreGetDoc("Usuarios", uid);
    if (!admin) return { ok: false, message: 'Administrador no encontrado en Firebase.' };
    
    return {
        ok: true,
        rol: 'admin',
        nombre: String(admin['NOMBRE'] || '').trim(),
        uid: uid
    };
}

// ─── Gestión de Usuarios (solo rol 'usuario') ─────────────────────────────────

/**
 * Crea un nuevo usuario con rol 'usuario' en la hoja config.
 * Solo pueden ejecutarlo admin o supervisor.
 * Estructura hoja config: ID | NOMBRE | USUARIO | PASS | ROL
 */
function crearUsuarioCD(nombreRaw: string, usuarioRaw: string, pass: string, solicitanteRol: string, autorCod: string): { ok: boolean; message?: string } {
  try {
    const nombre = sanitizeInput(nombreRaw);
    const usuario = sanitizeInput(usuarioRaw);
    if (solicitanteRol !== 'admin' && solicitanteRol !== 'supervisor') {
      return { ok: false, message: 'Sin permisos para crear usuarios.' };
    }
    
    const todosUsuarios = firestoreGetAllDocs("Usuarios");
    const usrInput = usuario.trim().toLowerCase();

    for (const u of todosUsuarios) {
      if (String(u['USUARIO'] || '').trim().toLowerCase() === usrInput) {
        return { ok: false, message: 'El nombre de usuario "' + usuario + '" ya existe.' };
      }
    }

    let maxId = 0;
    for (const u of todosUsuarios) {
      const idNum = parseInt(String(u['ID_USER'] || '0'), 10);
      if (!isNaN(idNum) && idNum > maxId) maxId = idNum;
    }
    const newId = maxId + 1;

    const rowObject = {
      'ID_USER': newId,
      'NOMBRE': nombre.trim(),
      'USUARIO': usuario.trim(),
      'PASS': pass.trim(),
      'ROL': 'usuario',
      'AUTOR_COD': (autorCod || '').trim().toUpperCase(),
      'CORREO': '',
      'STATUS': true
    };

    firestoreUpdateDocument("Usuarios", String(newId), rowObject);
    return { ok: true, message: 'Usuario "' + usuario + '" creado exitosamente (ID: ' + newId + ').' };
  } catch (e) {
    return { ok: false, message: 'Error al crear usuario (Firebase): ' + e.message };
  }
}

/**
 * Lista todos los usuarios con rol 'usuario' (excluye admin y supervisor).
 * Solo pueden ejecutarlo admin o supervisor.
 */
function listarUsuariosCD(solicitanteRol: string): { ok: boolean; data?: any[]; message?: string } {
  try {
    if (solicitanteRol !== 'admin' && solicitanteRol !== 'supervisor') {
      return { ok: false, message: 'Sin permisos para ver usuarios.' };
    }
    const todosUsuarios = firestoreGetAllDocs("Usuarios");
    const result = todosUsuarios.filter(u => String(u['ROL']).toLowerCase() === 'usuario');

    return { ok: true, data: result };
  } catch (e) {
    return { ok: false, message: 'Error al listar usuarios (Firebase): ' + e.message };
  }
}

/**
 * Cambia el estado de inhabilitación (Columna F) de un usuario en la hoja config.
 */
function cambiarStatusUsuarioCD(idUsuario: string, nuevoStatus: boolean, solicitanteRol: string): { ok: boolean; message?: string } {
  try {
    if (solicitanteRol !== 'admin' && solicitanteRol !== 'supervisor') {
      return { ok: false, message: 'Sin permisos para editar usuarios.' };
    }
    
    const user = firestoreGetDoc("Usuarios", idUsuario);
    if (!user) return { ok: false, message: 'Usuario no encontrado.' };

    firestoreUpdateDocument("Usuarios", idUsuario, { 'STATUS': nuevoStatus });
    
    const stStr = nuevoStatus ? 'habilitado' : 'inhabilitado';
    Logger.log('Usuario ID ' + idUsuario + ' ' + stStr + ' en Firebase');
    return { ok: true, message: 'Usuario ' + stStr + ' exitosamente.' };
  } catch (e) {
    return { ok: false, message: 'Error al cambiar status (Firebase): ' + e.message };
  }
}


/**
 * Elimina un usuario por ID. Solo pueden ejecutarlo admin o supervisor.
 * No permite eliminar usuarios con rol admin o supervisor.
 */
function eliminarUsuarioCD(idUsuario: string, solicitanteRol: string): { ok: boolean; message?: string } {
  try {
    if (solicitanteRol !== 'admin' && solicitanteRol !== 'supervisor') {
      return { ok: false, message: 'Sin permisos para eliminar usuarios.' };
    }

    const user = firestoreGetDoc("Usuarios", idUsuario);
    if (!user) return { ok: false, message: 'Usuario con ID ' + idUsuario + ' no encontrado.' };

    const rolUsuario = String(user['ROL'] || '').trim().toLowerCase();
    if (rolUsuario === 'admin' || rolUsuario === 'supervisor') {
      return { ok: false, message: 'No se puede eliminar un usuario con rol admin o supervisor.' };
    }

    // --- ELIMINACION DE CHAT PRIVADO ---
    try {
      const sp = PropertiesService.getScriptProperties();
      sp.deleteProperty('CHAT_MSGS_IND_' + idUsuario);
      const cache = CacheService.getScriptCache();
      cache.remove('CHAT_MSGS_CACHE_' + idUsuario);
    } catch(e) { Logger.log("Error limpiando chat: " + e.message); }

    firestoreDeleteDocument("Usuarios", idUsuario);
    Logger.log('Usuario eliminado de Firebase: ID=' + idUsuario);

    return { ok: true, message: 'Usuario y su historial de chat eliminados de Firebase.' };
  } catch (e) {
    return { ok: false, message: 'Error al eliminar usuario (Firebase): ' + e.message };
  }
}

// ─── Inicialización de hojas ───────────────────────────────────────────────────

/** Crea las hojas REQUERIMIENTOS_CD y DOCUMENTOS_CD si no existen. */
function initHojasCD(): { ok: boolean; message: string } {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // REQUERIMIENTOS_CD
    if (!ss.getSheetByName(CD_HOJA_REQUERIMIENTOS)) {
      const h = ss.insertSheet(CD_HOJA_REQUERIMIENTOS);
      h.getRange(1, 1, 1, 8).setValues([[
        'ID_REQ', 'ID_USUARIO', 'APELLIDO_NOMBRE', 'AUTORIDAD',
        'FECHA_APERTURA', 'ESTADO', 'ID_PDF_DRIVE', 'ID_CARPETA_DRIVE'
      ]]);
    }

    // DOCUMENTOS_CD
    if (!ss.getSheetByName(CD_HOJA_DOCUMENTOS)) {
      const h = ss.insertSheet(CD_HOJA_DOCUMENTOS);
      h.getRange(1, 1, 1, 9).setValues([[
        'ID_DOC', 'ID_REQ', 'NUM_REQ', 'ESTADO', 'URL_DRIVE',
        'FECHA_CARGA', 'FECHA_APROBACION', 'OBS_RECHAZO', 'APROBADO_POR'
      ]]);
    }

    return { ok: true, message: 'Hojas inicializadas correctamente.' };
  } catch (e) {
    return { ok: false, message: 'Error al inicializar hojas: ' + e.message };
  }
}

// ─── Días hábiles ─────────────────────────────────────────────────────────────

/** Calcula cuántos días hábiles faltan desde hoy hasta la fecha de vencimiento (apertura + 7 hábiles). */
function calcularVencimientoCD(fechaAperturaStr: string): { diasRestantes: number; vencido: boolean; fechaVencimiento: string } {
  const apertura = new Date(fechaAperturaStr);
  let diasHabiles = 0;
  let fecha = new Date(apertura);
  while (diasHabiles < 7) {
    fecha.setDate(fecha.getDate() + 1);
    const dow = fecha.getDay();
    if (dow !== 0 && dow !== 6) diasHabiles++; // excluir sáb y dom
  }
  const vencimiento = new Date(fecha);
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  vencimiento.setHours(0, 0, 0, 0);

  // Calcular días hábiles restantes entre hoy y vencimiento
  let restantes = 0;
  const cursor = new Date(hoy);
  if (hoy <= vencimiento) {
    while (cursor <= vencimiento) {
      const d = cursor.getDay();
      if (d !== 0 && d !== 6) restantes++;
      cursor.setDate(cursor.getDate() + 1);
    }
  }

  const fv = Utilities.formatDate(vencimiento, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  return {
    diasRestantes: restantes,
    vencido: hoy > vencimiento,
    fechaVencimiento: fv
  };
}

// ─── Requerimientos ───────────────────────────────────────────────────────────

/** Crea un nuevo requerimiento con los 15 documentos en PENDIENTE. */
function crearRequerimiento(apellidoNombreRaw: string, autoridadRaw: string, usuarioId: string): { ok: boolean; message?: string; idReq?: string } {
  try {
    const uid = usuarioId;
    const apellidoNombre = sanitizeInput(apellidoNombreRaw).toUpperCase();
    const autoridad = sanitizeInput(autoridadRaw);
    if (!uid) return { ok: false, message: 'Identificador de usuario faltante.' };

    // Generar ID_REQ basado en el último correlativo
    const todosReqs = firestoreGetAllDocs("Requerimientos");
    let maxId = 0;
    for (const r of todosReqs) {
      const idNum = parseInt(String(r['ID_REQ'] || '0'), 10);
      if (!isNaN(idNum) && idNum > maxId) maxId = idNum;
    }
    const newId = maxId + 1;
    const idReq = String(newId);

    const fechaApertura = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    
    // Crear subcarpeta en DRIVE
    const folderName = `${apellidoNombre.trim()}_${autoridad.trim()}`;
    const mainFolder = DriveApp.getFolderById(CARPETA_DOCUMENTACION);
    const subFolder = mainFolder.createFolder(folderName);
    subFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const idCarpetaDrive = subFolder.getId();

    // Guardar en Firestore: Requerimientos
    const reqData = {
      'ID_REQ': idReq,
      'ID_USUARIO': uid,
      'APELLIDO_NOMBRE': apellidoNombre.trim(),
      'AUTORIDAD': autoridad.trim(),
      'FECHA_APERTURA': fechaApertura,
      'ESTADO': 'EN CURSO',
      'ID_PDF_DRIVE': '',
      'ID_CARPETA_DRIVE': idCarpetaDrive
    };
    firestoreUpdateDocument("Requerimientos", idReq, reqData);

    // Guardar en Firestore: 15 Documentos PENDIENTES
    const todosDocs = firestoreGetAllDocs("Documentos");
    let lastDocId = 0;
    for (const d of todosDocs) {
      const idNum = parseInt(String(d['DOC_ID'] || '0'), 10);
      if (!isNaN(idNum) && idNum > lastDocId) lastDocId = idNum;
    }

    for (let i = 1; i <= 15; i++) {
      const docId = String(lastDocId + i);
      const docData = {
        'DOC_ID': docId,
        'ID_REQ': idReq,
        'NUM_REQ': i,
        'ESTADO': 'PENDIENTE',
        'URL_DRIVE': '',
        'FECHA_CARGA': '',
        'FECHA_APROBACION': '',
        'OBS_RECHAZO': '',
        'APROBADO_POR': ''
      };
      firestoreUpdateDocument("Documentos", docId, docData);
    }

    return { ok: true, idReq: idReq };
  } catch (e) {
    return { ok: false, message: 'Error al crear requerimiento (Firebase): ' + e.message };
  }
}

/** Lista requerimientos. El usuario ve los suyos; el supervisor/admin ve todos. */
function listarRequerimientos(usuarioId: string, rol: string): { ok: boolean; data?: any[]; message?: string } {
  try {
    const uid   = usuarioId;
    const rол   = (rol || '').toLowerCase();
    if (!uid) return { ok: false, message: 'Datos de sesión incompletos.' };

    const todosReqs = firestoreGetAllDocs("Requerimientos");
    const todosDocs = firestoreGetAllDocs("Documentos");
    const todosUsuarios = firestoreGetAllDocs("Usuarios");

    // Map de aprobados: ID_REQ -> cantidad
    const aprobadosMap: {[key: string]: number} = {};
    for (const d of todosDocs) {
      if (d['ESTADO'] === 'APROBADO') {
        const idR = String(d['ID_REQ']);
        aprobadosMap[idR] = (aprobadosMap[idR] || 0) + 1;
      }
    }

    // Map de usuarios: ID_USER -> NOMBRE
    const usuariosMap: {[key: string]: string} = {};
    for (const u of todosUsuarios) {
      usuariosMap[String(u['ID_USER'])] = String(u['NOMBRE']);
    }

    const result: any[] = [];
    for (const row of todosReqs) {
      const idReq = row['ID_REQ'];
      const idUsuario = row['ID_USUARIO'];
      const apellidoNombre = row['APELLIDO_NOMBRE'];
      const autoridad = row['AUTORIDAD'];
      const fechaApertura = row['FECHA_APERTURA'];
      const estado = row['ESTADO'];
      
      // Filtrado por rol
      if ((rол === 'usuario') && String(idUsuario) !== uid) continue;

      const idReqStr = String(idReq);
      const aprobados = aprobadosMap[idReqStr] || 0;
      const nombreUsuario = usuariosMap[String(idUsuario)] || String(idUsuario);

      const venc = fechaApertura ? calcularVencimientoCD(String(fechaApertura)) : { diasRestantes: 0, vencido: false, fechaVencimiento: '' };

      result.push({
        idReq: idReqStr,
        idUsuario: String(idUsuario),
        nombreUsuario,
        apellidoNombre: String(apellidoNombre),
        autoridad: String(autoridad),
        fechaApertura: String(fechaApertura),
        estado: String(estado),
        idPdfDrive: String(row['ID_PDF_DRIVE'] || ''),
        idCarpetaDrive: String(row['ID_CARPETA_DRIVE'] || ''),
        aprobados,
        total: 15,
        diasRestantes: venc.diasRestantes,
        vencido: venc.vencido,
        fechaVencimiento: venc.fechaVencimiento
      });
    }

    return { ok: true, data: result };
  } catch (e) {
    return { ok: false, message: 'Error al listar requerimientos (Firebase): ' + e.message };
  }
}

/** Retorna los 15 documentos de un requerimiento específico. */
function getDocumentosReq(idReq: string): { ok: boolean; data?: any[]; message?: string } {
  try {
    const todosDocs = firestoreGetAllDocs("Documentos");
    const docs = todosDocs
      .filter(d => String(d['ID_REQ']) === String(idReq))
      .map(d => ({
        idDoc:          String(d['__id']),
        idReq:          String(d['ID_REQ']),
        numReq:         Number(d['NUM_REQ']),
        nombreReq:      CD_REQUISITOS[Number(d['NUM_REQ']) - 1] || 'Requisito ' + d['NUM_REQ'],
        estado:         String(d['ESTADO']),
        urlDrive:       String(d['URL_DRIVE'] || ''),
        fechaCarga:     String(d['FECHA_CARGA'] || ''),
        fechaAprobacion: String(d['FECHA_APROBACION'] || ''),
        obsRechazo:     String(d['OBS_RECHAZO'] || ''),
        aprobadoPor:    String(d['APROBADO_POR'] || '')
      }))
      .sort((a, b) => a.numReq - b.numReq);

    return { ok: true, data: docs };
  } catch (e) {
    return { ok: false, message: 'Error al obtener documentos (Firebase): ' + e.message };
  }
}

// ─── Documentos ────────────────────────────────────────────────────────────────

/** Actualiza una fila de DOCUMENTOS_CD. Helper interno. */


/** El usuario sube un archivo local (codificado en base64) a Google Drive y registra la URL en el sistema. */
function subirDocumentoLocalCD(idDoc: string, mimeType: string, base64Data: string, fileName: string): { ok: boolean; message?: string } {
  try {
    const doc = firestoreGetDoc("Documentos", idDoc);
    if (!doc) return { ok: false, message: 'Documento no encontrado en Firebase.' };

    const idReq = String(doc['ID_REQ']);
    const req = firestoreGetDoc("Requerimientos", idReq);
    let idCarpeta = req ? String(req['ID_CARPETA_DRIVE'] || CARPETA_DOCUMENTACION) : CARPETA_DOCUMENTACION;

    // Decodificar el base64 a un Blob
    const dataContent = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(dataContent, mimeType, fileName);

    const carpeta = DriveApp.getFolderById(idCarpeta);
    const file = carpeta.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const urlDrive = file.getUrl();

    const fechaCarga = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    
    // Actualizar Firestore
    const updateData = {
      'ESTADO': 'CARGADO',
      'URL_DRIVE': urlDrive,
      'FECHA_CARGA': fechaCarga,
      'OBS_RECHAZO': ''
    };
    firestoreUpdateDocument("Documentos", idDoc, updateData);

    return { ok: true };
  } catch (e) {
    return { ok: false, message: 'Error al subir archivo (Firebase): ' + e.message };
  }
}

/** El supervisor aprueba un documento. */
function aprobarDocumento(idDoc: string, nombreAprobador: string): { ok: boolean; message?: string } {
  try {
    const fechaAprobacion = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    
    // 1. Obtención del documento de Firestore
    const doc = firestoreGetDoc("Documentos", idDoc);
    if (!doc) return { ok: false, message: 'Documento no encontrado en Firebase.' };

    // 2. Actualización en Firestore
    const updateData = {
      'ESTADO': 'APROBADO',
      'FECHA_APROBACION': fechaAprobacion,
      'OBS_RECHAZO': '',
      'APROBADO_POR': nombreAprobador
    };
    firestoreUpdateDocument("Documentos", idDoc, updateData);

    // 3. Chequear si el Requerimiento se completó
    _checkCompletitudReqFirestore(String(doc['ID_REQ']));
    
    return { ok: true };
  } catch (e) {
    return { ok: false, message: 'Error al aprobar (Firebase): ' + e.message };
  }
}

/** Aprueba múltiples documentos en una sola operación (Simulado por ahora con bucle) */
function aprobarDocumentosLote(idDocs: string[], nombreAprobador: string): { ok: boolean; message?: string } {
  try {
    const fechaAprobacion = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    const reqsParaChequear = new Set<string>();

    for (const idDoc of idDocs) {
      const doc = firestoreGetDoc("Documentos", idDoc);
      if (doc) {
        firestoreUpdateDocument("Documentos", idDoc, {
          'ESTADO': 'APROBADO',
          'FECHA_APROBACION': fechaAprobacion,
          'OBS_RECHAZO': '',
          'APROBADO_POR': nombreAprobador
        });
        reqsParaChequear.add(String(doc['ID_REQ']));
      }
    }

    // Chequear todos los requerimientos afectados
    for (const idReq of reqsParaChequear) {
      _checkCompletitudReqFirestore(idReq);
    }

    return { ok: true };
  } catch (e) {
    return { ok: false, message: 'Error en aprobación masiva (Firebase): ' + e.message };
  }
}

/** Verifica en Firestore si un Requerimiento tiene sus 15 documentos aprobados. */
function _checkCompletitudReqFirestore(idReq: string): void {
  try {
    const todosDocs = firestoreGetAllDocs("Documentos");
    const docsDelReq = todosDocs.filter(d => String(d['ID_REQ']) === idReq);
    
    const aprobados = docsDelReq.filter(d => d['ESTADO'] === 'APROBADO').length;
    
    if (aprobados === 15) {
      firestoreUpdateDocument("Requerimientos", idReq, { 'ESTADO': 'COMPLETO' });
    }
  } catch (e) {
    Logger.log("Error en _checkCompletitudReqFirestore: " + e.message);
  }
}

/** El supervisor rechaza un documento con un motivo y elimina el archivo físico de Drive. */
function rechazarDocumento(idDoc: string, motivoRaw: string): { ok: boolean; message?: string } {
  try {
    const motivo = sanitizeInput(motivoRaw);
    const doc = firestoreGetDoc("Documentos", idDoc);
    if (!doc) return { ok: false, message: 'Documento no encontrado en Firebase.' };

    const urlParaBorrar = String(doc['URL_DRIVE'] || '');

    // Borrado físico en Drive
    if (urlParaBorrar) {
       try {
          const fileIdMatch = urlParaBorrar.match(/[-\w]{25,}/);
          if (fileIdMatch) {
             DriveApp.getFileById(fileIdMatch[0]).setTrashed(true);
          }
       } catch (e) {
          Logger.log("No se pudo borrar el archivo en Drive al rechazar: " + e.message);
       }
    }

    const updateData = {
      'ESTADO': 'RECHAZADO',
      'URL_DRIVE': '',
      'FECHA_APROBACION': '',
      'OBS_RECHAZO': motivo,
      'APROBADO_POR': ''
    };
    firestoreUpdateDocument("Documentos", idDoc, updateData);

    return { ok: true };
  } catch (e) {
    return { ok: false, message: 'Error al rechazar (Firebase): ' + e.message };
  }
}

/** El supervisor elimina un requerimiento completo y todos sus documentos. */
function eliminarRequerimientoComp(idReq: string, uid: string, rol: string): { ok: boolean; message?: string } {
  try {
    if (rol !== 'supervisor' && rol !== 'admin') {
      return { ok: false, message: 'No tienes permisos para eliminar.' };
    }
    
    const req = firestoreGetDoc("Requerimientos", idReq);
    if (!req) return { ok: false, message: 'Requerimiento no encontrado en Firebase.' };

    const idCarpetaTrash = String(req['ID_CARPETA_DRIVE'] || '');

    // Eliminar documentos vinculados
    const todosDocs = firestoreGetAllDocs("Documentos");
    for (const d of todosDocs) {
      if (String(d['ID_REQ']) === idReq) {
        firestoreDeleteDocument("Documentos", String(d['__id']));
      }
    }

    // Eliminar requerimiento
    firestoreDeleteDocument("Requerimientos", idReq);

    // Borrar carpeta en Drive
    if (idCarpetaTrash) {
      try {
        DriveApp.getFolderById(idCarpetaTrash).setTrashed(true);
      } catch (e) {
        Logger.log("Error al borrar carpeta en Drive: " + e.message);
      }
    }

    return { ok: true, message: 'Requerimiento y carpeta eliminados de Firebase y Drive.' };
  } catch (e) {
    return { ok: false, message: 'Error al eliminar requerimiento (Firebase): ' + e.message };
  }
}

/** El usuario elimina un documento rechazado para poder subir uno nuevo. */
function eliminarDocumento(idDoc: string): { ok: boolean; message?: string } {
  try {
    const updateData = {
      'ESTADO': 'PENDIENTE',
      'URL_DRIVE': '',
      'FECHA_CARGA': '',
      'OBS_RECHAZO': ''
    };
    firestoreUpdateDocument("Documentos", idDoc, updateData);
    return { ok: true };
  } catch (e) {
    return { ok: false, message: 'Error al eliminar (Firebase): ' + e.message };
  }
}

// ─── Generación de PDF ────────────────────────────────────────────────────────

/**
 * Genera el PDF de checklist aprobado para un requerimiento.
 * Retorna la URL del PDF en Drive.
 */
async function generarPdfCD(idReq: string, nombreAprobador: string): Promise<{ ok: boolean; urlPdf?: string; message?: string }> {
  try {
    // 1. Obtener datos del requerimiento de Firestore
    const reqData = firestoreGetDoc("Requerimientos", idReq);
    if (!reqData) return { ok: false, message: 'Requerimiento no encontrado en Firebase.' };

    // 2. Obtener documentos aprobados
    const docsResult = getDocumentosReq(idReq);
    if (!docsResult.ok || !docsResult.data) return { ok: false, message: 'Error al obtener documentos.' };
    const docs = docsResult.data.filter(d => d.estado === 'APROBADO');
    
    if (docs.length === 0) {
      return { ok: false, message: 'Debe haber al menos un documento aprobado para generar el PDF.' };
    }

    // 3. Preparar PDF
    let idCarpetaFinal = String(reqData['ID_CARPETA_DRIVE'] || CARPETA_DOCUMENTACION);
    const titulo = `CHECKLIST_${String(reqData['APELLIDO_NOMBRE']).replace(/,/g, '').replace(/ /g, '_')}_${idReq}`;
    const doc = DocumentApp.create(titulo);
    const body = doc.getBody();

    body.appendParagraph('CHECKLIST - CONTRATO DE LOCACIÓN')
      .setHeading(DocumentApp.ParagraphHeading.HEADING1)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('');
    body.appendParagraph(`Contratado: ${reqData['APELLIDO_NOMBRE']}`).editAsText().setBold(true);
    body.appendParagraph(`Autoridad Solicitante: ${reqData['AUTORIDAD']}`);
    body.appendParagraph(`Fecha de apertura: ${reqData['FECHA_APERTURA']}`);
    body.appendParagraph(`Supervisor que aprobó: ${nombreAprobador || 'No indicado'}`);
    body.appendParagraph(`Fecha de generación: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')}`);
    body.appendHorizontalRule();

    const pdfBlobs: GoogleAppsScript.Base.Blob[] = [];
    doc.saveAndClose();
    const caratulaBlob = DriveApp.getFileById(doc.getId()).getBlob().setName('00_Reporte.pdf');
    pdfBlobs.push(caratulaBlob);
    
    for (const d of docs) {
      if (d.urlDrive) {
        try {
          const match = d.urlDrive.match(/[-\w]{25,}/);
          if (match) pdfBlobs.push(DriveApp.getFileById(match[0]).getBlob());
        } catch (e) {}
      }
    }

    let finalBlob: GoogleAppsScript.Base.Blob;
    try {
      finalBlob = await PDFApp.mergePDFs(pdfBlobs);
      finalBlob.setName(titulo + '.pdf');
    } catch (e) {
      finalBlob = caratulaBlob.setName(titulo + '.pdf');
    }

    const pdfFile = DriveApp.getFolderById(idCarpetaFinal).createFile(finalBlob);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    DriveApp.getFileById(doc.getId()).setTrashed(true);

    const urlPdf = pdfFile.getUrl();

    // Actualizar Firestore
    firestoreUpdateDocument("Requerimientos", idReq, {
      'ID_PDF_DRIVE': pdfFile.getId(),
      'ESTADO': 'PDF_GENERADO'
    });

    return { ok: true, urlPdf };
  } catch (e) {
    return { ok: false, message: 'Error al generar PDF (Firebase): ' + e.message };
  }
}

/** Retorna la URL de previsualización del PDF ya generado. */
function getPdfUrlCD(idReq: string): { ok: boolean; urlPdf?: string; message?: string } {
  try {
    const req = firestoreGetDoc("Requerimientos", idReq);
    if (!req) return { ok: false, message: 'Requerimiento no encontrado en Firebase.' };

    if (req['ID_PDF_DRIVE']) {
      const fileId = String(req['ID_PDF_DRIVE']);
      const url = `https://drive.google.com/file/d/${fileId}/view`;
      return { ok: true, urlPdf: url };
    }
    
    return { ok: false, message: 'PDF no generado aún.' };
  } catch (e) {
    return { ok: false, message: 'Error (Firebase): ' + e.message };
  }
}

/** Asegura que un archivo sea público para visualización rápida antes de abrir el visor. */
function asegurarAccesoPublico(url: string): { ok: boolean } {
  try {
    const match = url.match(/[-\w]{25,}/);
    if (match) {
      const file = DriveApp.getFileById(match[0]);
      // Solo aplicar si no tiene ya el permiso correcto (para ahorrar tiempo)
      if (file.getSharingAccess() !== DriveApp.Access.ANYONE_WITH_LINK) {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      }
    }
    return { ok: true };
  } catch (e) {
    return { ok: false };
  }
}

/** URL de la página de Control Docus (para el botón del Index principal). */
function getControlDocusUrl(): string {
  return ScriptApp.getService().getUrl() + '?page=PanelSupervisor';
}

// ─── Chat Global de Soporte ────────────────────────────────────────────────────

/**
 * Sincroniza el chat: reporta conexión y recupera mensajes nuevos.
 */
function chatSyncCD(uid: string, rol: string, nombre: string, lastMsgId: number, ghostMode: boolean = false, targetUid?: string): { ok: boolean; msgs: any[]; onlineInfo: any; error?: string } {
  try {
    const sp = PropertiesService.getScriptProperties();
    const cache = CacheService.getScriptCache();
    const now = Date.now();
    
    const esUsuario = (rol.toLowerCase() === 'usuario');
    const threadUid = esUsuario ? uid : targetUid;

    // 1. Reportar presencia con throttling
    const userLastSyncKey = "LAST_SYNC_" + uid;
    const lastUserSync = cache.get(userLastSyncKey);
    
    let onlineDict: any = {};
    const onlineStr = sp.getProperty('CHAT_ONLINE') || '{}';
    try { onlineDict = JSON.parse(onlineStr); } catch (e) {}

    if (!lastUserSync || (now - Number(lastUserSync) > 30000)) {
      onlineDict[uid] = { rol: rol.toLowerCase(), nombre: nombre, ts: now, ghost: ghostMode };
      for (let key in onlineDict) {
        if (now - onlineDict[key].ts > 120000) delete onlineDict[key];
      }
      sp.setProperty('CHAT_ONLINE', JSON.stringify(onlineDict));
      cache.put(userLastSyncKey, String(now), 60);
    }
    
    // 2. Preparar info de quiénes están online
    let activeUsuarios: any = {};
    let activeSupervisores: any = {};
    for (let key in onlineDict) {
      const u = onlineDict[key];
      // Para saber cuándo expira visualmente o ocultos
      if (!u.ghost && (now - u.ts < 90000)) {
        if (u.rol === 'usuario') activeUsuarios[key] = u.nombre;
        else activeSupervisores[key] = u.nombre;
      }
    }
    
    // 3. Recuperar mensajes del HILO ESPECIFICO (solo si hay threadUid)
    let nuevosMsgs: any[] = [];
    if (threadUid) {
      const CACHE_MSGS_KEY = "CHAT_MSGS_CACHE_" + threadUid;
      const PROP_MSGS_KEY = "CHAT_MSGS_IND_" + threadUid;
      let msgsArr: any[] = [];
      const cachedMsgs = cache.get(CACHE_MSGS_KEY);
      
      if (cachedMsgs) {
        msgsArr = JSON.parse(cachedMsgs);
      } else {
        const msgsStr = sp.getProperty(PROP_MSGS_KEY) || '[]';
        try { msgsArr = JSON.parse(msgsStr); } catch (e) {}
        cache.put(CACHE_MSGS_KEY, JSON.stringify(msgsArr), 15); // Cache por 15 segundos para hit rate alto
      }
      
      nuevosMsgs = msgsArr.filter((m: any) => m.id > lastMsgId);
    }
    
    return { 
      ok: true, 
      msgs: nuevosMsgs, 
      onlineInfo: { usuarios: activeUsuarios, supervisores: activeSupervisores }
    };
  } catch (e) {
    return { ok: false, msgs: [], onlineInfo: {}, error: e.message };
  }
}


/**
 * Envía un mensaje al chat global.
 */
function chatSendMsgCD(uid: string, rol: string, nombre: string, text: string, ghostMode: boolean = false, targetUid?: string): { ok: boolean; error?: string } {
  try {
    const threadUid = (rol.toLowerCase() === 'usuario') ? uid : targetUid;
    if (!threadUid) return { ok: false, error: "No se especificó destinatario" };

    const txtLim = sanitizeInput(text).substring(0, 500);
    if (!txtLim.trim()) return { ok: true };
    const sp = PropertiesService.getScriptProperties();
    const PROP_MSGS_KEY = "CHAT_MSGS_IND_" + threadUid;
    
    let msgsStr = sp.getProperty(PROP_MSGS_KEY) || '[]';
    let msgsArr: any[] = [];
    try { msgsArr = JSON.parse(msgsStr); } catch(e){}
    
    let nextId = 1;
    if (msgsArr.length > 0) nextId = msgsArr[msgsArr.length - 1].id + 1;
    
    const nuevoMsg = {
      id: nextId,
      uid: uid,
      rol: rol,
      nombre: nombre,
      text: txtLim,
      ts: Date.now()
    };
    
    msgsArr.push(nuevoMsg);
    if(msgsArr.length > 60) msgsArr = msgsArr.slice(-60);
    
    sp.setProperty(PROP_MSGS_KEY, JSON.stringify(msgsArr));
    
    // Invalidar cache
    const cache = CacheService.getScriptCache();
    cache.remove("CHAT_MSGS_CACHE_" + threadUid);

    chatSyncCD(uid, rol, nombre, nextId, ghostMode, targetUid); 
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}
