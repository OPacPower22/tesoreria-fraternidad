const SPREADSHEET_ID = "1gUULgTmfFbBsPK5DDaV8TX7Q5Cnv-vCXCVxeXO3xAeM";
const ADMIN_EMAIL = "fraternidad.num1@gmail.com";

/*************************************************
 UTILIDAD BASE
*************************************************/

function obtenerSpreadsheet(){
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function esAdmin(){
  const email = Session.getActiveUser().getEmail().toLowerCase().trim();
  return email === ADMIN_EMAIL.toLowerCase().trim();
}

/*************************************************
 WEB APP
*************************************************/

function doGet(){
  return HtmlService
    .createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Tesorería Fraternidad No. 1");
}

function incluir(nombre){
  return HtmlService.createHtmlOutputFromFile(nombre).getContent();
}

function obtenerVista(nombre){
  return HtmlService.createHtmlOutputFromFile(nombre).getContent();
}

/*************************************************
 FECHA Y HORA
*************************************************/

function obtenerFechaHoraServidor(){

  const ahora = new Date();
  const zona = Session.getScriptTimeZone();

  const dias = ["domingo","lunes","martes","miércoles","jueves","viernes","sábado"];
  const meses = ["enero","febrero","marzo","abril","mayo","junio",
                 "julio","agosto","septiembre","octubre","noviembre","diciembre"];

  return {
    fecha: `${dias[ahora.getDay()]}, ${ahora.getDate()} DE ${meses[ahora.getMonth()]} DE ${ahora.getFullYear()}`.toUpperCase(),
    hora: Utilities.formatDate(ahora,zona,"HH:mm:ss") + " HRS."
  };
}

/*************************************************
 DASHBOARD
*************************************************/

function obtenerResumenMensual(){

  const meses = [
    "enero","febrero","marzo","abril","mayo","junio",
    "julio","agosto","septiembre","octubre","noviembre","diciembre"
  ];

  const mesActual = meses[new Date().getMonth()];

  const hoja = obtenerSpreadsheet().getSheetByName("CONTROL_MENSUAL");
  const datos = hoja.getDataRange().getValues();

  for(let i=1;i<datos.length;i++){

    const mes = (datos[i][0]||"").toLowerCase().trim();

    if(mes === mesActual){

      return {
        mesActual: mesActual,
        ingresos: Number(datos[i][3]||0),
        egresos: Number(datos[i][4]||0),
        saldo: Number(datos[i][5]||0)
      };

    }
  }

  return {
    mesActual,
    ingresos:0,
    egresos:0,
    saldo:0
  };
}

function obtenerEstadoMesActual(){

  const meses = ["enero","febrero","marzo","abril","mayo","junio",
                 "julio","agosto","septiembre","octubre","noviembre","diciembre"];

  const mesActual = meses[new Date().getMonth()];

  const hoja = obtenerSpreadsheet().getSheetByName("CONTROL_MENSUAL");
  const datos = hoja.getDataRange().getValues();

  for(let i=1;i<datos.length;i++){
    if((datos[i][0]||"").toLowerCase().trim() === mesActual){
      return { estado: datos[i][2] };
    }
  }

  return { estado:"ABIERTO" };
}

function obtenerDashboardCompleto(){
  return {
    fechaHora: obtenerFechaHoraServidor(),
    resumen: obtenerResumenMensual(),
    estado: obtenerEstadoMesActual()
  };
}

/*************************************************
 CONTROL MENSUAL
*************************************************/

function mesEstaCerrado(mes){

  const hoja = obtenerSpreadsheet().getSheetByName("CONTROL_MENSUAL");
  const datos = hoja.getDataRange().getValues();

  mes = mes.toLowerCase().trim();

  for(let i=1;i<datos.length;i++){
    const mesControl = (datos[i][0]||"").toLowerCase().trim();
    const estado = datos[i][2];
    if(mesControl === mes && estado === "CERRADO"){
      return true;
    }
  }

  return false;
}

/*************************************************
 REGISTRAR MOVIMIENTO
*************************************************/

function registrarMovimiento(datos){

  if(!datos) throw new Error("Datos no recibidos.");

  const mes = datos.mes.toLowerCase().trim();

// Si no es mes actual
if(!esMesActual(mes)){

  // Solo permitir si es admin
  if(!esAdmin()){
    throw new Error("MES_NO_PERMITIDO");
  }

  // Si es admin pero el mes está cerrado, exigir autorización
  if(mesEstaCerrado(mes)){
    throw new Error("MES_CERRADO");
  }
}


  const hoja = obtenerSpreadsheet().getSheetByName("REGISTRO_GENERAL");

  const ultimaFila = hoja.getLastRow();
  const nuevoID = ultimaFila>1
    ? hoja.getRange(ultimaFila,1).getValue()+1
    : 1;

  const monto = Number((datos.monto||"").toString().replace(/[^0-9.-]+/g,""));
  if(isNaN(monto)||monto<=0) throw new Error("Monto inválido.");

  hoja.appendRow([
    nuevoID,
    new Date(),
    datos.mes.toLowerCase().trim(),
    new Date().getFullYear(),
    datos.tipo.toLowerCase().trim(),
    datos.etiqueta||"",
    datos.miembro||"",
    "",
    monto,
    datos.metodoPago||"",
    datos.referencia||"",
    new Date()
  ]);

  actualizarSaldoMes(datos.mes, datos.tipo.toLowerCase().trim(), monto);

  const props = PropertiesService.getScriptProperties();

const mesTemporal = props.getProperty("MES_REABIERTO_TEMP");
const operacionPendiente = props.getProperty("OPERACION_REAPERTURA_PENDIENTE");

if(mesTemporal && mesTemporal === mes && operacionPendiente === "true"){

  const hojaControl = obtenerSpreadsheet().getSheetByName("CONTROL_MENSUAL");
  const datosControl = hojaControl.getDataRange().getValues();

  for(let i=1;i<datosControl.length;i++){

    if((datosControl[i][0]||"").toLowerCase().trim() === mes){

      hojaControl.getRange(i+1,3).setValue("CERRADO");
      hojaControl.getRange(i+1,7).setValue(new Date());

      break;
    }

  }

  props.deleteProperty("MES_REABIERTO_TEMP");
  props.deleteProperty("OPERACION_REAPERTURA_PENDIENTE");

  registrarBitacora(
    "CIERRE_AUTOMATICO_MES",
    "Mes: "+mes+" | cierre automático tras reapertura"
  );

}

  registrarBitacora("REGISTRO_MOVIMIENTO","ID:"+nuevoID+" | Mes:"+datos.mes);

  return {status:"ok",id:nuevoID};
}

/*************************************************
 BITÁCORA
*************************************************/

function registrarBitacora(accion,detalle){

  const hoja = obtenerSpreadsheet().getSheetByName("BITACORA_AUDITORIA");

  const ultimaFila = hoja.getLastRow();
  const nuevoID = ultimaFila>1
    ? hoja.getRange(ultimaFila,1).getValue()+1
    : 1;

  hoja.appendRow([
    nuevoID,
    new Date(),
    Session.getActiveUser().getEmail(),
    esAdmin()?"ADMIN":"CONSULTA",
    accion,
    detalle
  ]);
}

/*************************************************
 2FA
*************************************************/

function validarTOTP(codigo){

  if(!codigo) return false;

  // asegurar que sea un código de 6 dígitos
  codigo = codigo.toString().trim();
  if(!/^\d{6}$/.test(codigo)) return false;

  const props = PropertiesService.getScriptProperties();
  const secret = props.getProperty("ADMIN_TOTP_SECRET");

  if(!secret) return false;

  const timeStep = Math.floor(Date.now()/1000/30);

  // ventana de tolerancia ±2 intervalos
  for(let i=-2;i<=2;i++){
    const generado = generarTOTP(secret,timeStep+i);
    if(generado === codigo){
      return true;
    }
  }

  return false;
}

/*************************************************
 USUARIO ACTUAL
*************************************************/

function obtenerUsuarioActual(){

  const email = Session.getActiveUser().getEmail();

  if(!email){
    throw new Error("No se pudo obtener el correo del usuario.");
  }

  const rol = email.toLowerCase().trim() === ADMIN_EMAIL.toLowerCase().trim()
    ? "ADMIN"
    : "CONSULTA";

  return {
    email: email,
    rol: rol
  };
}

/*************************************************
 MES ACTUAL
*************************************************/

function esMesActual(mes){

  const meses = [
    "enero","febrero","marzo","abril","mayo","junio",
    "julio","agosto","septiembre","octubre","noviembre","diciembre"
  ];

  const mesActual = meses[new Date().getMonth()];

  return mes.toLowerCase().trim() === mesActual;
}

/*************************************************
 FUNCION PARA CONSULTAR ESTADO ACTUAL 2FA
*************************************************/

function estado2FA(){

  const props = PropertiesService.getScriptProperties();
  const activado = props.getProperty("ADMIN_2FA_ACTIVADO");

  return {
    activado: activado === "true"
  };
}

/*************************************************
 FUNCION PARA GENERAR QR SOLO SI NO ESTA ACTIVADO
*************************************************/

function generarQR2FA(){

  const props = PropertiesService.getScriptProperties();
  const activado = props.getProperty("ADMIN_2FA_ACTIVADO");

  if(activado === "true"){
    return { mostrarQR:false };
  }

  const secret = props.getProperty("ADMIN_TOTP_SECRET");
  const email = props.getProperty("ADMIN_EMAIL");

  const issuer = "FraternidadNo1";

  const uri =
    "otpauth://totp/" +
    encodeURIComponent(issuer + ":" + email) +
    "?secret=" + secret +
    "&issuer=" + issuer +
    "&algorithm=SHA1&digits=6&period=30";

  return {
    mostrarQR:true,
    qr:
      "https://api.qrserver.com/v1/create-qr-code/?size=250x250&data=" +
      encodeURIComponent(uri)
  };
}

/*************************************************
 FUNCION PARA CONFIRMAR ACTIVACION
*************************************************/
function confirmarActivacion2FA(codigo){

  if(!validarTOTP(codigo)){
    throw new Error("CODIGO_INVALIDO");
  }

  const props = PropertiesService.getScriptProperties();
  props.setProperty("ADMIN_2FA_ACTIVADO","true");

  return true;
}

/*************************************************
 AUTORIZAR REAPERTURA CON CREDENCIALES
*************************************************/

function autorizarReapertura(mes,email,password,codigo){

  const props = PropertiesService.getScriptProperties();

  const adminEmail = ADMIN_EMAIL.toLowerCase().trim();
  const hashGuardado = props.getProperty("ADMIN_HASH");

  const hashIngresado = Utilities.base64Encode(
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      password
    )
  );

  Logger.log("Email ingresado: " + email);
  Logger.log("Hash ingresado: " + hashIngresado);
  Logger.log("Hash guardado: " + hashGuardado);

  if(email.trim().toLowerCase() !== adminEmail){
    throw new Error("NO_AUTORIZADO");
  }

  if(hashIngresado !== hashGuardado){
    throw new Error("CREDENCIALES_INVALIDAS");
  }

  if(!validarTOTP(codigo)){
    throw new Error("2FA_INVALIDO");
  }

  props.setProperty("ADMIN_2FA_ACTIVADO","true");
  props.setProperty("MES_REABIERTO_TEMP", mes.toLowerCase());
  props.setProperty("OPERACION_REAPERTURA_PENDIENTE","true");

  const hoja = obtenerSpreadsheet().getSheetByName("CONTROL_MENSUAL");
  const datos = hoja.getDataRange().getValues();

  for(let i=1;i<datos.length;i++){
    if((datos[i][0]||"").toLowerCase().trim() === mes.toLowerCase()){
      hoja.getRange(i+1,3).setValue("ABIERTO");
      break;
    }
  }

  registrarBitacora(
    "REAPERTURA_MES",
    "Mes: "+mes+" | Autorizado con 2FA"
  );

  return true;
}

/***********************************
DROP LIST DE MIEMBROS
************************************/

function obtenerListaMiembros(){

  const hoja = obtenerSpreadsheet().getSheetByName("MIEMBROS");
  const datos = hoja.getDataRange().getValues();

  const encabezados = datos[0];
  const indexNombre = encabezados.indexOf("Nombre");

  if(indexNombre === -1){
    throw new Error("Columna 'Nombre' no encontrada en MIEMBROS.");
  }

  const lista = [];

  for(let i=1;i<datos.length;i++){
    const nombre = (datos[i][indexNombre] || "").toString().trim();
    if(nombre){
      lista.push(nombre);
    }
  }

  return lista.sort();
}

/***********************************
CONTRASEÑA
************************************/

/*function generarHashPrueba(){
  const password = "FRATERNIDAD2026";
  const hash = Utilities.base64Encode(
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      password
    )
  );
  Logger.log(hash);
}*/

/***********************************
TOTP TEMPORAL
************************************/

function debugTOTP(){
  const props = PropertiesService.getScriptProperties();
  const secret = props.getProperty("ADMIN_TOTP_SECRET");
  const timeStep = Math.floor(Date.now()/1000/30);
  Logger.log(generarTOTP(secret,timeStep));
}

/***********************************
NUEVO SECRETO TOTP TEMPORAL
************************************/

function generarNuevoSecretTOTP() {

  const caracteres = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567";
  let secret = "";

  for (let i = 0; i < 16; i++) {
    const randomIndex = Math.floor(Math.random() * caracteres.length);
    secret += caracteres[randomIndex];
  }

  Logger.log("Nuevo SECRET:");
  Logger.log(secret);
}

/***********************************
RESET DE 2FA CONTROLADO
************************************/

/*function reset2FA(){

  const props = PropertiesService.getScriptProperties();

  props.setProperty("ADMIN_2FA_ACTIVADO","false");

  Logger.log("2FA reiniciado");
}

function reset2FACompleto(){

  const props = PropertiesService.getScriptProperties();

  const caracteres="ABCDEFGHIJKLMNOPQRSTUVWXYZ234567";
  let secret="";

  for(let i=0;i<16;i++){
    secret+=caracteres[Math.floor(Math.random()*caracteres.length)];
  }

  props.setProperty("ADMIN_TOTP_SECRET",secret);
  props.setProperty("ADMIN_2FA_ACTIVADO","false");

  Logger.log("Nuevo SECRET:");
  Logger.log(secret);
}*/

/*************************************************
AUTENTICACION ADMIN
*************************************************/

function autenticarAdmin(email,password){

  const props = PropertiesService.getScriptProperties();
  const hashGuardado = props.getProperty("ADMIN_HASH");

  const hashIngresado = Utilities.base64Encode(
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      password
    )
  );

  if(email.trim().toLowerCase() !== ADMIN_EMAIL.toLowerCase()){
    throw new Error("USUARIO_INVALIDO");
  }

  if(hashIngresado !== hashGuardado){
    throw new Error("PASSWORD_INVALIDO");
  }

  props.setProperty("ADMIN_AUTH_TEMP","true");

  return true;
}

/*************************************************
VALIDACION 2FA
*************************************************/

function validarCodigo2FA(codigo){

  const props = PropertiesService.getScriptProperties();
  const auth = props.getProperty("ADMIN_AUTH_TEMP");

  if(auth !== "true"){
    throw new Error("NO_AUTENTICADO");
  }

  if(!validarTOTP(codigo)){
    throw new Error("2FA_INVALIDO");
  }

  props.setProperty("ADMIN_2FA_OK","true");

  return true;
}

/*************************************************
REAPERTURA SEGURA
*************************************************/

function reabrirMesSeguro(mes){

  const props = PropertiesService.getScriptProperties();
  const autorizado = props.getProperty("ADMIN_2FA_OK");

  if(autorizado !== "true"){
    throw new Error("NO_AUTORIZADO");
  }

  const hoja = obtenerSpreadsheet().getSheetByName("CONTROL_MENSUAL");
  const datos = hoja.getDataRange().getValues();

  for(let i=1;i<datos.length;i++){
    if((datos[i][0]||"").toLowerCase().trim() === mes.toLowerCase()){
      hoja.getRange(i+1,3).setValue("ABIERTO");
      break;
    }
  }

  registrarBitacora(
    "REAPERTURA_MES",
    "Mes: "+mes+" | Autorización completa"
  );

  props.deleteProperty("ADMIN_AUTH_TEMP");
  props.deleteProperty("ADMIN_2FA_OK");

  return true;
}

/*************************************************
GENERAR TOTP
*************************************************/

function generarTOTP(secret, timeStep) {

  const key = base32Decode(secret);

  const signature = Utilities.computeHmacSignature(
    Utilities.MacAlgorithm.HMAC_SHA_1,
    intToBytes(timeStep),
    key
  );

  const offset = signature[signature.length - 1] & 0xf;

  const binary =
    ((signature[offset] & 0x7f) << 24) |
    ((signature[offset + 1] & 0xff) << 16) |
    ((signature[offset + 2] & 0xff) << 8) |
    (signature[offset + 3] & 0xff);

  const otp = binary % 1000000;

  return otp.toString().padStart(6, "0");
}

function base32Decode(input){

  const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567";
  let bits = "";
  let value = 0;
  let index = 0;
  const output = [];

  input = input.replace(/=+$/,"");

  for (let i = 0; i < input.length; i++) {

    value = (value << 5) | alphabet.indexOf(input[i].toUpperCase());
    index += 5;

    if (index >= 8) {
      output.push((value >>> (index - 8)) & 255);
      index -= 8;
    }

  }

  return output;
}

function intToBytes(num){

  const bytes = [];

  for (let i = 7; i >= 0; i--) {
    bytes[i] = num & 0xff;
    num = num >> 8;
  }

  return bytes;
}

/*************************************************
GOOGLE AUTH
*************************************************/

function debugTOTP(){

  const props = PropertiesService.getScriptProperties();
  const secret = props.getProperty("ADMIN_TOTP_SECRET");

  const timeStep = Math.floor(Date.now()/1000/30);

  const codigo = generarTOTP(secret,timeStep);

  Logger.log("Codigo actual servidor:");
  Logger.log(codigo);
}

/*************************************************
CERRAR EL MES REABIERTO
*************************************************/

function cerrarMes(mes){

  const hoja = obtenerSpreadsheet().getSheetByName("CONTROL_MENSUAL");
  const datos = hoja.getDataRange().getValues();

  for(let i=1;i<datos.length;i++){

    if((datos[i][0]||"").toLowerCase().trim() === mes.toLowerCase()){

      hoja.getRange(i+1,3).setValue("CERRADO");

      registrarBitacora(
        "CIERRE_MES",
        "Mes: "+mes
      );

      return true;
    }

  }

  throw new Error("MES_NO_ENCONTRADO");
}

/*************************************************
ACTUALIZAR SALDO DEL MES
*************************************************/

function actualizarSaldoMes(mes, tipo, monto){

  const hoja = obtenerSpreadsheet().getSheetByName("CONTROL_MENSUAL");
  const datos = hoja.getDataRange().getValues();

  mes = mes.toLowerCase().trim();

  for(let i=1;i<datos.length;i++){

    const mesControl = (datos[i][0]||"").toLowerCase().trim();

    if(mesControl === mes){

      let ingresos = Number(datos[i][3]||0);
      let egresos = Number(datos[i][4]||0);

      tipo = tipo.toLowerCase().trim();

    if(tipo === "ingreso"){
      ingresos += monto;
    }

    if(tipo === "egreso"){
      egresos += monto;
    }

      const saldo = ingresos - egresos;

      hoja.getRange(i+1,4).setValue(ingresos);
      hoja.getRange(i+1,5).setValue(egresos);
      hoja.getRange(i+1,6).setValue(saldo);

      return;
    }
  }
}
