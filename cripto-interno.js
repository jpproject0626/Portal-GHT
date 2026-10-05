/* Descifrado en el navegador de excluidos_avance.json (Web Crypto).
   El archivo lo cifra generar_datos.py con AES-256-GCM; la llave se deriva de
   la clave interna con PBKDF2-SHA256 usando el salt y las iteraciones que trae
   el propio archivo. Aqui NO hay ninguna clave escrita: descifrar con exito es
   lo que prueba que la clave es la correcta. */
(function () {
  const ARCHIVO = 'excluidos_avance.json';
  const MIN_ITER = 100000, MAX_ITER = 5000000;

  // Error con un codigo para que quien llama muestre el mensaje adecuado:
  // 'red' (no se pudo bajar el archivo), 'formato' (no es un sobre cifrado
  // valido), 'clave' (la clave no descifra), 'cripto' (el navegador no
  // soporta Web Crypto).
  function errorCripto(codigo, mensaje) {
    const e = new Error(mensaje);
    e.codigo = codigo;
    return e;
  }

  function base64ABytes(b64) {
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
  }

  async function descifrarExcluidos(clave) {
    if (!(window.crypto && crypto.subtle)) {
      throw errorCripto('cripto', 'Este navegador no soporta Web Crypto.');
    }

    let sobre;
    try {
      const resp = await fetch(ARCHIVO + '?v=' + Date.now(), { cache: 'no-store' });
      if (!resp.ok) throw new Error('HTTP ' + resp.status);
      sobre = await resp.json();
    } catch (err) {
      throw errorCripto('red', 'No se pudo leer ' + ARCHIVO + ': ' + err.message);
    }

    let salt, iv, cifrado;
    try {
      if (!sobre || sobre.v !== 1 || sobre.alg !== 'AES-256-GCM' || sobre.kdf !== 'PBKDF2-SHA256' ||
          !(sobre.iter >= MIN_ITER && sobre.iter <= MAX_ITER)) {
        throw new Error('parametros no validos');
      }
      salt = base64ABytes(sobre.salt);
      iv = base64ABytes(sobre.iv);
      cifrado = base64ABytes(sobre.ct);
    } catch (err) {
      throw errorCripto('formato', 'El archivo no tiene el formato cifrado esperado.');
    }

    const material = await crypto.subtle.importKey(
      'raw', new TextEncoder().encode(String(clave).trim()), 'PBKDF2', false, ['deriveKey']);
    const llave = await crypto.subtle.deriveKey(
      { name: 'PBKDF2', salt, iterations: sobre.iter, hash: 'SHA-256' },
      material, { name: 'AES-GCM', length: 256 }, false, ['decrypt']);

    let claro;
    try {
      claro = await crypto.subtle.decrypt({ name: 'AES-GCM', iv }, llave, cifrado);
    } catch (err) {
      throw errorCripto('clave', 'La clave no descifra el archivo.');
    }
    try {
      return JSON.parse(new TextDecoder().decode(claro));
    } catch (err) {
      throw errorCripto('formato', 'El contenido descifrado no es valido.');
    }
  }

  window.descifrarExcluidos = descifrarExcluidos;
})();
