/* app.js - versión ajustada
   - Fecha en formato argentino (DD/MM/AAAA) dentro del PDF
   - Año dinámico en "ORDEN DE SERVICIO Nro"
   - "Tipo de Vuelo" resaltado en una línea aparte
   - Títulos claros para: Personal Terrestre, Personal de Seguridad y Vehículos
*/

(function () {
  // Helpers básicos
  function qs(id) { return document.getElementById(id); }
  function $all(sel, ctx = document) { return Array.from(ctx.querySelectorAll(sel)); }

  function loadScript(src) {
    return new Promise((resolve, reject) => {
      if (document.querySelector(`script[src="${src}"]`)) return resolve();
      const s = document.createElement('script');
      s.src = src;
      s.onload = () => resolve();
      s.onerror = () => reject(new Error('Error loading ' + src));
      document.head.appendChild(s);
    });
  }

  // localStorage seguro
  function safeSetItem(k, v) {
    try { localStorage.setItem(k, v); return true; } catch (e) { console.warn('localStorage.setItem failed', e); return false; }
  }
  function safeGetItem(k) {
    try { return localStorage.getItem(k); } catch (e) { console.warn('localStorage.getItem failed', e); return null; }
  }

  // Cargar una imagen y convertirla a dataURL
  function loadImageAsDataURL(src) {
    return new Promise((resolve) => {
      const img = new Image();
      img.crossOrigin = 'Anonymous';
      img.onload = () => {
        try {
          const canvas = document.createElement('canvas');
          canvas.width = img.naturalWidth;
          canvas.height = img.naturalHeight;
          const ctx = canvas.getContext('2d');
          ctx.drawImage(img, 0, 0);
          const dataURL = canvas.toDataURL('image/png');
          resolve(dataURL);
        } catch (e) {
          console.warn('convert image failed', e);
          resolve(null);
        }
      };
      img.onerror = () => resolve(null);
      img.src = src;
    });
  }

  // 👉 NUEVO: formatear fecha YYYY-MM-DD a DD/MM/YYYY
  function formatearFechaArg(fechaIso) {
    if (!fechaIso) return '';
    const partes = fechaIso.split('-');
    if (partes.length !== 3) return fechaIso;
    const [anio, mes, dia] = partes;
    if (!anio || !mes || !dia) return fechaIso;
    return `${dia}/${mes}/${anio}`;
  }

  // Elementos
  const btnHoraInicio = qs('btnHoraInicio');
  const btnHoraFin = qs('btnHoraFin');
  const btnHoraPartida = qs('btnHoraPartida');
  const btnHoraArribo = qs('btnHoraArribo');
  const btnAddPT = qs('btnAddPT');
  const btnAddPS = qs('btnAddPS');
  const btnAddVh = qs('btnAddVh');
  const btnClear = qs('btnClear');
  const btnPreview = qs('btnPreview');
  const btnPreviewBack = qs('btnPreviewBack');
  const btnPreviewConfirm = qs('btnPreviewConfirm');
  const previewModal = qs('previewModal');
  const signatureModal = qs('signatureModal');
  const closePreview = qs('closePreview');
  const closeSignature = qs('closeSignature');
  const btnConfirmarFirma = qs('btnConfirmarFirma');
  const btnClearSig = qs('btnClearSig');
  const signatureCanvas = qs('signatureCanvas');
  const autosaveIndicator = qs('autosaveIndicator');
  const errorContainer = qs('errorContainer');

  const personalTerrestreTable = qs('personalTerrestreTable');
  const personalSeguridadTable = qs('personalSeguridadTable');
  const vehiculosContainer = qs('vehiculosContainer');

  // Estado firma
  let canvasCtx = null;
  let drawing = false;
  let hasFirmado = false;
  let firmaImagen = null;
  let signatureInitialized = false;
  let autosaveInterval = null;

  // INIT
  function init() {
    agregarPersonalTerrestre();
    agregarPersonalSeguridad();
    agregarVehiculo();

    const fechaControl = qs('fechaControl');
    if (fechaControl) fechaControl.valueAsDate = new Date();

    // Botones de hora
    btnHoraInicio && btnHoraInicio.addEventListener('click', () => setCurrentTime('horaInicio'));
    btnHoraFin && btnHoraFin.addEventListener('click', () => setCurrentTime('horaFin'));
    btnHoraPartida && btnHoraPartida.addEventListener('click', () => setCurrentTime('horaPartida'));
    btnHoraArribo && btnHoraArribo.addEventListener('click', () => setCurrentTime('horaArribo'));

    // Radio visual
    document.querySelectorAll('.radio-option').forEach(opt => {
      opt.addEventListener('click', () => {
        const groupParent = opt.parentElement;
        if (groupParent) Array.from(groupParent.children).forEach(c => c.classList.remove('selected'));
        opt.classList.add('selected');
        const input = opt.querySelector('input');
        if (input) input.checked = true;
        if (input && input.name === 'tipoVuelo') cambiarTipoVuelo(input.value);
        if (input && input.name === 'cantPersonal') selectPersonal(input.value);
      });
    });

    // Botones de agregar / limpiar
    btnAddPT && btnAddPT.addEventListener('click', agregarPersonalTerrestre);
    btnAddPS && btnAddPS.addEventListener('click', agregarPersonalSeguridad);
    btnAddVh && btnAddVh.addEventListener('click', agregarVehiculo);
    btnClear && btnClear.addEventListener('click', limpiarFormulario);

    // Preview / firma
    btnPreview && btnPreview.addEventListener('click', abrirVistaPrevia);
    closePreview && closePreview.addEventListener('click', cerrarVistaPrevia);
    btnPreviewBack && btnPreviewBack.addEventListener('click', cerrarVistaPrevia);
    btnPreviewConfirm && btnPreviewConfirm.addEventListener('click', confirmarYFirmar);

    // Modal firma
    closeSignature && closeSignature.addEventListener('click', cerrarModalFirma);
    btnClearSig && btnClearSig.addEventListener('click', limpiarFirmaModal);
    btnConfirmarFirma && btnConfirmarFirma.addEventListener('click', confirmarFirma);

    // Autosave
    startAutosave();

    // Preview live
    document.getElementById('planillaForm').addEventListener('input', updatePreview);

    updatePreview();
  }

  function setCurrentTime(fieldId) {
    const now = new Date();
    const hours = String(now.getHours()).padStart(2, '0');
    const minutes = String(now.getMinutes()).padStart(2, '0');
    const el = qs(fieldId);
    if (el) {
      el.value = `${hours}:${minutes}`;
      updatePreview();
    }
  }

  function selectPersonal(value) {
    const otroField = qs('cantPersonalOtro');
    if (value === 'otro') {
      if (otroField) otroField.style.display = 'block';
    } else {
      if (otroField) {
        otroField.style.display = 'none';
        otroField.value = '';
      }
    }
    updatePreview();
  }

  function cambiarTipoVuelo(tipo) {
    const origen = qs('origen');
    const destino = qs('destino');
    const campoPartida = qs('campoHoraPartida');
    const campoArribo = qs('campoHoraArribo');

    if (tipo === 'Salida') {
      if (origen) { origen.value = 'AEP'; origen.disabled = true; }
      if (destino) destino.disabled = false;
      if (campoPartida) campoPartida.style.display = 'flex';
      if (campoArribo) campoArribo.style.display = 'none';
      qs('horaArribo') && (qs('horaArribo').value = '');
    } else {
      if (destino) { destino.value = 'AEP'; destino.disabled = true; }
      if (origen) origen.disabled = false;
      if (campoPartida) campoPartida.style.display = 'none';
      if (campoArribo) campoArribo.style.display = 'flex';
      qs('horaPartida') && (qs('horaPartida').value = '');
    }
    updatePreview();
  }

  // Tablas dinámicas
  function agregarPersonalTerrestre() {
    const table = personalTerrestreTable;
    const row = table.insertRow();
    row.innerHTML = `
      <td><input type="text" name="pt_nombre[]" class="form-input" placeholder="PEREZ Juan"></td>
      <td><input type="text" name="pt_dni[]" class="form-input" placeholder="12345678"></td>
      <td><input type="text" name="pt_legajo[]" class="form-input" placeholder="12345"></td>
      <td>
        <select name="pt_funcion[]" class="form-select">
          <option value="">-</option>
          <option value="Sup">Sup</option>
          <option value="Bod">Bod</option>
          <option value="Cin">Cin</option>
          <option value="Tra">Tra</option>
          <option value="Otr">Otr</option>
        </select>
      </td>
      <td><input type="text" name="pt_grupo[]" class="form-input" placeholder="Grupo A"></td>
      <td><button type="button" class="btn btn-remove">✕</button></td>
    `;
    row.querySelector('.btn-remove').addEventListener('click', () => { row.remove(); updatePreview(); });
    updatePreview();
  }

  function agregarPersonalSeguridad() {
    const table = personalSeguridadTable;
    const row = table.insertRow();
    row.innerHTML = `
      <td><input type="text" name="ps_nombre[]" class="form-input" placeholder="PEREZ Juan"></td>
      <td><input type="text" name="ps_dni[]" class="form-input" placeholder="12345678"></td>
      <td><input type="text" name="ps_legajo[]" class="form-input" placeholder="12345"></td>
      <td><input type="text" name="ps_empresa[]" class="form-input" placeholder="SegurSA"></td>
      <td><button type="button" class="btn btn-remove">✕</button></td>
    `;
    row.querySelector('.btn-remove').addEventListener('click', () => { row.remove(); updatePreview(); });
    updatePreview();
  }

  function agregarVehiculo() {
    const container = vehiculosContainer;
    const index = container.children.length;
    const div = document.createElement('div');
    div.className = 'vehicle-card';
    div.id = `vh${index}`;
    div.innerHTML = `
      <div class="form-grid form-grid-2">
        <div class="form-field">
          <label class="form-label">Tipo Vehículo</label>
          <input type="text" name="vh_tipo[]" class="form-input" placeholder="Tractor">
        </div>
        <div class="form-field">
          <label class="form-label">Empresa</label>
          <input type="text" name="vh_empresa[]" class="form-input" placeholder="Aerolíneas">
        </div>
        <div class="form-field">
          <label class="form-label">Nº Interno</label>
          <input type="text" name="vh_interno[]" class="form-input" placeholder="123">
        </div>
        <div class="form-field">
          <label class="form-label">Operador</label>
          <input type="text" name="vh_operador[]" class="form-input" placeholder="GOMEZ Pedro">
        </div>
      </div>
      <div class="form-field" style="margin-top:12px;">
        <label class="form-label">Observaciones</label>
        <input type="text" name="vh_obs[]" class="form-input" placeholder="Observaciones...">
      </div>
      <button type="button" class="btn btn-remove" style="margin-top:12px;">Eliminar Vehículo</button>
    `;
    div.querySelector('.btn-remove').addEventListener('click', () => { div.remove(); updatePreview(); });
    container.appendChild(div);
    updatePreview();
  }

  // Autosave
  function guardarBorrador() {
    const data = {};
    const campos = [
      'ordenServicio','fechaControl','patrulla','horaInicio','horaFin',
      'empresa','codigoVuelo','matricula','origen','destino','posicion',
      'horaArribo','horaPartida','conDemora',
      'novEquipajes','novInspeccion','novEquipajesUtil','otrasNovedades',
      'respNombre','respDNI','respLegajo'
    ];
    campos.forEach(campo => {
      const el = qs(campo);
      if (el) data[campo] = el.value;
    });

    data.tipoVuelo = document.querySelector('input[name="tipoVuelo"]:checked')?.value;
    data.cantPersonal = document.querySelector('input[name="cantPersonal"]:checked')?.value;
    data.cantPersonalOtro = qs('cantPersonalOtro').value || '';

    data.ctrlPersonas = !!qs('ctrlPersonas') && qs('ctrlPersonas').checked;
    data.ctrlEquipos  = !!qs('ctrlEquipos')  && qs('ctrlEquipos').checked;
    data.ctrlCargas   = !!qs('ctrlCargas')   && qs('ctrlCargas').checked;

    data.medioMovil   = !!qs('medioMovil')   && qs('medioMovil').checked;
    data.medioPaletas = !!qs('medioPaletas') && qs('medioPaletas').checked;
    data.medioOtros   = !!qs('medioOtros')   && qs('medioOtros').checked;

    data.tipoProcedimiento = document.querySelector('input[name="tipoProcedimiento"]:checked')?.value || '';

    data.personalTerrestre = [];
    $all('input[name="pt_nombre[]"]').forEach((el, i) => {
      data.personalTerrestre.push({
        nombre: el.value,
        dni:     ($all('input[name="pt_dni[]"]')[i]     || {}).value || '',
        legajo:  ($all('input[name="pt_legajo[]"]')[i]  || {}).value || '',
        funcion: ($all('select[name="pt_funcion[]"]')[i]|| {}).value || '',
        grupo:   ($all('input[name="pt_grupo[]"]')[i]   || {}).value || ''
      });
    });

    data.personalSeguridad = [];
    $all('input[name="ps_nombre[]"]').forEach((el, i) => {
      data.personalSeguridad.push({
        nombre: el.value,
        dni:     ($all('input[name="ps_dni[]"]')[i]     || {}).value || '',
        legajo:  ($all('input[name="ps_legajo[]"]')[i]  || {}).value || '',
        empresa: ($all('input[name="ps_empresa[]"]')[i] || {}).value || ''
      });
    });

    data.vehiculos = [];
    $all('input[name="vh_tipo[]"]').forEach((el, i) => {
      data.vehiculos.push({
        tipo:     el.value,
        empresa:  ($all('input[name="vh_empresa[]"]')[i] || {}).value || '',
        interno:  ($all('input[name="vh_interno[]"]')[i] || {}).value || '',
        operador: ($all('input[name="vh_operador[]"]')[i]|| {}).value || '',
        obs:      ($all('input[name="vh_obs[]"]')[i]     || {}).value || ''
      });
    });

    safeSetItem('borradorPlanilla', JSON.stringify(data));
    showAutosaveIndicator();
  }

  function startAutosave() {
    if (autosaveInterval) clearInterval(autosaveInterval);
    autosaveInterval = setInterval(guardarBorrador, 30000);
  }

  function showAutosaveIndicator() {
    autosaveIndicator.classList.add('show');
    setTimeout(() => autosaveIndicator.classList.remove('show'), 1500);
  }

  // Vista previa
  function updatePreview() {
    const preview = qs('previewContent');
    if (!preview) return;

    let cantPersonal = document.querySelector('input[name="cantPersonal"]:checked')?.value || '-';
    if (cantPersonal === 'otro') cantPersonal = qs('cantPersonalOtro').value || 'Especificar';

    const tipoVuelo = document.querySelector('input[name="tipoVuelo"]:checked')?.value || '-';
    const tipoProcedimiento = document.querySelector('input[name="tipoProcedimiento"]:checked')?.value || '-';

    const tiposControl = [];
    if (qs('ctrlPersonas') && qs('ctrlPersonas').checked) tiposControl.push('Personas');
    if (qs('ctrlEquipos') && qs('ctrlEquipos').checked) tiposControl.push('Equipos');
    if (qs('ctrlCargas') && qs('ctrlCargas').checked) tiposControl.push('Cargas');

    const medios = [];
    if (qs('medioMovil') && qs('medioMovil').checked) medios.push('Móvil');
    if (qs('medioPaletas') && qs('medioPaletas').checked) medios.push('Paletas');
    if (qs('medioOtros') && qs('medioOtros').checked) medios.push('Otros');

    const nombresT   = $all('input[name="pt_nombre[]"]').map(i => i.value || '—');
    const dnisT      = $all('input[name="pt_dni[]"]').map(i => i.value || '—');
    const legajosT   = $all('input[name="pt_legajo[]"]').map(i => i.value || '—');
    const funcionesT = $all('select[name="pt_funcion[]"]').map(i => i.value || '-');
    const gruposT    = $all('input[name="pt_grupo[]"]').map(i => i.value || '-');

    let personalTerrestreHTML = '';
    for (let i = 0; i < nombresT.length; i++) {
      personalTerrestreHTML += `
        <div class="preview-item">
          <span class="preview-label">${nombresT[i]}</span>
          <span class="preview-value">
            DNI: ${dnisT[i]} — Legajo: ${legajosT[i]} — Función: ${funcionesT[i]} — Grupo: ${gruposT[i]}
          </span>
        </div>`;
    }
    if (!personalTerrestreHTML) {
      personalTerrestreHTML = `
        <div class="preview-item">
          <span class="preview-label">Sin registros</span>
          <span class="preview-value">-</span>
        </div>`;
    }

    const nombresS   = $all('input[name="ps_nombre[]"]').map(i => i.value || '—');
    const dnisS      = $all('input[name="ps_dni[]"]').map(i => i.value || '—');
    const legajosS   = $all('input[name="ps_legajo[]"]').map(i => i.value || '—');
    const empresasS  = $all('input[name="ps_empresa[]"]').map(i => i.value || '-');

    let personalSeguridadHTML = '';
    for (let i = 0; i < nombresS.length; i++) {
      personalSeguridadHTML += `
        <div class="preview-item">
          <span class="preview-label">${nombresS[i]}</span>
          <span class="preview-value">
            DNI: ${dnisS[i]} — Legajo: ${legajosS[i]} — Empresa: ${empresasS[i]}
          </span>
        </div>`;
    }
    if (!personalSeguridadHTML) {
      personalSeguridadHTML = `
        <div class="preview-item">
          <span class="preview-label">Sin registros</span>
          <span class="preview-value">-</span>
        </div>`;
    }

    const tiposVh     = $all('input[name="vh_tipo[]"]').map(i => i.value || '—');
    const empresasVh  = $all('input[name="vh_empresa[]"]').map(i => i.value || '—');
    const internosVh  = $all('input[name="vh_interno[]"]').map(i => i.value || '—');
    const operadoresVh= $all('input[name="vh_operador[]"]').map(i => i.value || '—');
    const obsVh       = $all('input[name="vh_obs[]"]').map(i => i.value || '-');

    let vehiculosHTML = '';
    for (let i = 0; i < tiposVh.length; i++) {
      vehiculosHTML += `
        <div class="preview-item">
          <span class="preview-label">${tiposVh[i]}</span>
          <span class="preview-value">
            Empresa: ${empresasVh[i]} — Interno: ${internosVh[i]} — Operador: ${operadoresVh[i]} — Obs: ${obsVh[i]}
          </span>
        </div>`;
    }
    if (!vehiculosHTML) {
      vehiculosHTML = `
        <div class="preview-item">
          <span class="preview-label">Sin registros</span>
          <span class="preview-value">-</span>
        </div>`;
    }

    const html = `
      <div class="preview-section">
        <h4>🔍 Control PSA</h4>
        <div class="preview-item"><span class="preview-label">Orden Servicio:</span><span class="preview-value">${qs('ordenServicio').value || '-'}</span></div>
        <div class="preview-item"><span class="preview-label">Fecha:</span><span class="preview-value">${formatearFechaArg(qs('fechaControl').value || '') || '-'}</span></div>
        <div class="preview-item"><span class="preview-label">Patrulla:</span><span class="preview-value">${qs('patrulla').value || '-'}</span></div>
        <div class="preview-item"><span class="preview-label">Hora Inicio:</span><span class="preview-value">${qs('horaInicio').value || '-'}</span></div>
        <div class="preview-item"><span class="preview-label">Hora Fin:</span><span class="preview-value">${qs('horaFin').value || '-'}</span></div>
        <div class="preview-item"><span class="preview-label">Personal:</span><span class="preview-value">${cantPersonal}</span></div>
        <div class="preview-item"><span class="preview-label">Tipos Control:</span><span class="preview-value">${tiposControl.join(', ') || 'Ninguno'}</span></div>
        <div class="preview-item"><span class="preview-label">Medios Técnicos:</span><span class="preview-value">${medios.join(', ') || 'Ninguno'}</span></div>
        <div class="preview-item"><span class="preview-label">Tipo Procedimiento:</span><span class="preview-value">${tipoProcedimiento}</span></div>
      </div>

      <div class="preview-section">
        <h4>✈️ Vuelo</h4>
        <div class="preview-item"><span class="preview-label">Tipo:</span><span class="preview-value">${tipoVuelo}</span></div>
        <div class="preview-item"><span class="preview-label">Empresa:</span><span class="preview-value">${qs('empresa').value || '-'}</span></div>
        <div class="preview-item"><span class="preview-label">Código:</span><span class="preview-value">${qs('codigoVuelo').value || '-'}</span></div>
        <div class="preview-item"><span class="preview-label">Matrícula:</span><span class="preview-value">${qs('matricula').value || '-'}</span></div>
        <div class="preview-item"><span class="preview-label">Origen:</span><span class="preview-value">${qs('origen').value || '-'}</span></div>
        <div class="preview-item"><span class="preview-label">Destino:</span><span class="preview-value">${qs('destino').value || '-'}</span></div>
        <div class="preview-item"><span class="preview-label">Hora Partida/Arribo:</span><span class="preview-value">${qs('horaPartida').value || qs('horaArribo').value || '-'}</span></div>
        <div class="preview-item"><span class="preview-label">Con Demora:</span><span class="preview-value">${qs('conDemora').value}</span></div>
        <div class="preview-item"><span class="preview-label">Posición:</span><span class="preview-value">${qs('posicion').value || '-'}</span></div>
      </div>

      <div class="preview-section">
        <h4>👤 Responsable PSA</h4>
        <div class="preview-item"><span class="preview-label">Nombre:</span><span class="preview-value">${qs('respNombre').value || '-'}</span></div>
        <div class="preview-item"><span class="preview-label">DNI:</span><span class="preview-value">${qs('respDNI').value || '-'}</span></div>
        <div class="preview-item"><span class="preview-label">Legajo:</span><span class="preview-value">${qs('respLegajo').value || '-'}</span></div>
      </div>

      <div class="preview-section">
        <h4>👷 Personal de Apoyo Terrestre</h4>
        ${personalTerrestreHTML}
      </div>

      <div class="preview-section">
        <h4>🛡️ Personal de Seguridad</h4>
        ${personalSeguridadHTML}
      </div>

      <div class="preview-section">
        <h4>🚗 Vehículos Controlados</h4>
        ${vehiculosHTML}
      </div>

      <div class="preview-section">
        <h4>📝 Novedades</h4>
        <div class="preview-item"><span class="preview-label">Equipajes/Cargas:</span><span class="preview-value">${qs('novEquipajes').value || 'Sin novedades'}</span></div>
        <div class="preview-item"><span class="preview-label">Inspección del Personal:</span><span class="preview-value">${qs('novInspeccion').value || 'Sin novedades'}</span></div>
        <div class="preview-item"><span class="preview-label">Equipajes Utilizados:</span><span class="preview-value">${qs('novEquipajesUtil').value || 'Sin novedades'}</span></div>
        <div class="preview-item"><span class="preview-label">Otras:</span><span class="preview-value">${qs('otrasNovedades').value || 'Sin novedades'}</span></div>
      </div>
    `;
    preview.innerHTML = html;
  }

  function abrirVistaPrevia() {
    errorContainer.innerHTML = '';
    const required = [
      { id: 'respNombre', name: 'Nombre del Responsable PSA' },
      { id: 'respDNI', name: 'DNI del Responsable PSA' },
      { id: 'respLegajo', name: 'Legajo del Responsable PSA' },
      { id: 'fechaControl', name: 'Fecha de Control' },
      { id: 'horaInicio', name: 'Hora de Inicio' },
      { id: 'horaFin', name: 'Hora de Finalización' },
      { id: 'empresa', name: 'Empresa del Vuelo' },
      { id: 'codigoVuelo', name: 'Código de Vuelo' }
    ];

    const errores = [];
    required.forEach(c => {
      const el = qs(c.id);
      if (!el || !el.value || !el.value.toString().trim()) {
        errores.push(`Falta completar: ${c.name}`);
        el && el.classList.add('error');
      } else el.classList.remove('error');
    });

    if (errores.length) {
      let html = '<div class="alert alert-danger"><strong>⚠️ Completá los siguientes campos antes de continuar:</strong><ul>';
      errores.forEach(e => html += `<li>${e}</li>`);
      html += '</ul></div>';
      errorContainer.innerHTML = html;
      return;
    }

    updatePreview();
    previewModal.style.display = 'block';
    document.body.style.overflow = 'hidden';
  }

  function cerrarVistaPrevia() {
    previewModal.style.display = 'none';
    document.body.style.overflow = 'auto';
  }

  function confirmarYFirmar() {
    cerrarVistaPrevia();
    abrirModalFirma();
  }

  function abrirModalFirma() {
    signatureModal.style.display = 'block';
    document.body.style.overflow = 'hidden';
    if (!signatureInitialized) initSignatureCanvas();
    limpiarFirmaModal();
  }

  function cerrarModalFirma() {
    signatureModal.style.display = 'none';
    document.body.style.overflow = 'auto';
  }

  function initSignatureCanvas() {
    canvasCtx = signatureCanvas.getContext('2d');
    canvasCtx.strokeStyle = '#000';
    canvasCtx.lineWidth = 2;
    canvasCtx.lineCap = 'round';
    canvasCtx.lineJoin = 'round';

    signatureCanvas.addEventListener('mousedown', (e) => {
      drawing = true;
      const p = getCanvasPos(e);
      canvasCtx.beginPath();
      canvasCtx.moveTo(p.x, p.y);
    });
    signatureCanvas.addEventListener('mousemove', (e) => {
      if (!drawing) return;
      const p = getCanvasPos(e);
      canvasCtx.lineTo(p.x, p.y);
      canvasCtx.stroke();
      hasFirmado = true;
      btnConfirmarFirma.disabled = false;
    });
    signatureCanvas.addEventListener('mouseup', () => drawing = false);
    signatureCanvas.addEventListener('mouseout', () => drawing = false);

    signatureCanvas.addEventListener('touchstart', (e) => {
      e.preventDefault();
      drawing = true;
      const touch = e.touches[0];
      const p = getCanvasPos(touch);
      canvasCtx.beginPath();
      canvasCtx.moveTo(p.x, p.y);
    }, { passive: false });

    signatureCanvas.addEventListener('touchmove', (e) => {
      e.preventDefault();
      if (!drawing) return;
      const touch = e.touches[0];
      const p = getCanvasPos(touch);
      canvasCtx.lineTo(p.x, p.y);
      canvasCtx.stroke();
      hasFirmado = true;
      btnConfirmarFirma.disabled = false;
    }, { passive: false });

    signatureInitialized = true;
  }

  function getCanvasPos(e) {
    const rect = signatureCanvas.getBoundingClientRect();
    const scaleX = signatureCanvas.width / rect.width;
    const scaleY = signatureCanvas.height / rect.height;
    const clientX = e.clientX;
    const clientY = e.clientY;
    return {
      x: (clientX - rect.left) * scaleX,
      y: (clientY - rect.top) * scaleY
    };
  }

  function limpiarFirmaModal() {
    if (!canvasCtx) initSignatureCanvas();
    canvasCtx.clearRect(0, 0, signatureCanvas.width, signatureCanvas.height);
    hasFirmado = false;
    btnConfirmarFirma.disabled = true;
    firmaImagen = null;
  }

  function confirmarFirma() {
    if (!hasFirmado) {
      alert('Por favor, dibujá tu firma antes de confirmar');
      return;
    }
    firmaImagen = signatureCanvas.toDataURL('image/png');
    cerrarModalFirma();
    generarPDF();
  }

  function limpiarFormulario() {
    if (!confirm('¿Limpiar todo el formulario?')) return;
    qs('planillaForm').reset();
    personalTerrestreTable.innerHTML = '';
    personalSeguridadTable.innerHTML = '';
    vehiculosContainer.innerHTML = '';
    errorContainer.innerHTML = '';
    agregarPersonalTerrestre();
    agregarPersonalSeguridad();
    agregarVehiculo();
    cambiarTipoVuelo('Salida');
    firmaImagen = null;
    hasFirmado = false;
    updatePreview();
    safeSetItem('borradorPlanilla', '');
  }

  // ==============
  // GENERACIÓN PDF
  // ==============
  async function generarPDF() {
    try {
      if (!window.jspdf) {
        await loadScript('https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js');
      }
      if (!window.jspdf || !window.jspdf.jsPDF) throw new Error('jsPDF no disponible');

      if (!window.jspPDFAutoTableLoaded) {
        await loadScript('https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.28/jspdf.plugin.autotable.min.js');
        window.jspPDFAutoTableLoaded = true;
      }

      const jsPDF = window.jspdf.jsPDF;
      const doc = new jsPDF('p', 'mm', 'a4');
      const pageWidth = doc.internal.pageSize.getWidth();
      const pageHeight = doc.internal.pageSize.getHeight();
      const margin = 15;
      let y = 12;

      // Fecha cruda (input) y año
      const fechaIso = qs('fechaControl').value || '';
      const fechaArg = formatearFechaArg(fechaIso) || '-';
      const anio = fechaIso ? fechaIso.split('-')[0] : '';

      // Logos opcionales
      const logoLeft = await loadImageAsDataURL('logo_left.png');
      const logoRight = await loadImageAsDataURL('logo_right.png');
      const logoSize = 18;

      if (logoLeft) {
        try { doc.addImage(logoLeft, 'PNG', margin, y - 2, logoSize, logoSize); } catch (e) { }
      }
      if (logoRight) {
        try { doc.addImage(logoRight, 'PNG', pageWidth - margin - logoSize, y - 2, logoSize, logoSize); } catch (e) { }
      }

      // Título
      const title = 'PLANILLA DE CONTROL DE BODEGA UOSP METROPOLITANA';
      const titleHeight = 10;
      const titleY = y;
      const titleX = margin;
      const titleW = pageWidth - margin * 2;
      doc.setDrawColor(0);
      doc.setFillColor(255, 255, 255);
      doc.setLineWidth(0.8);
      doc.rect(titleX, titleY, titleW, titleHeight, 'FD');
      doc.setFontSize(12);
      doc.setFont(undefined, 'bold');
      doc.text(title, pageWidth / 2, titleY + 7, { align: 'center' });
      y += titleHeight + 6;

      // PATRULLA + ORDEN SERVICIO (con año dinámico)
      doc.setFontSize(9);
      doc.setFont(undefined, 'bold');
      doc.text('PATRULLA:', margin, y);
      doc.setFont(undefined, 'normal');
      doc.text(qs('patrulla').value || '-', margin + 28, y);

      doc.setFont(undefined, 'bold');
      doc.text('ORDEN DE SERVICIO Nro:', margin + 90, y);
      doc.setFont(undefined, 'normal');
      const ordenServ = qs('ordenServicio').value || '';
      doc.text(ordenServ, margin + 140, y);
      if (anio) {
        doc.text('/' + anio, margin + 140 + doc.getTextWidth(ordenServ) + 2, y);
      }
      y += 8;

      // CONTROL PSA
      const respNombre = qs('respNombre').value || '-';
      const horaInicio = qs('horaInicio').value || '-';
      const horaFin = qs('horaFin').value || '-';

      let cantidadPersonal = document.querySelector('input[name="cantPersonal"]:checked')?.value || '-';
      if (cantidadPersonal === 'otro') cantidadPersonal = qs('cantPersonalOtro').value || '-';

      const tiposControlArr = [];
      if (qs('ctrlPersonas') && qs('ctrlPersonas').checked) tiposControlArr.push('Personas');
      if (qs('ctrlEquipos') && qs('ctrlEquipos').checked) tiposControlArr.push('Equipos');
      if (qs('ctrlCargas') && qs('ctrlCargas').checked) tiposControlArr.push('Cargas');
      const tiposControlText = tiposControlArr.join(', ') || '-';

      const controlHead = [[
        'Fecha de Control',
        'Responsable PSA',
        'Hora de Inicio',
        'Hora de finalización',
        'Cantidad Personal',
        'Tipos de Controles'
      ]];
      const controlBody = [[
        fechaArg,
        respNombre,
        horaInicio,
        horaFin,
        cantidadPersonal,
        tiposControlText
      ]];

      doc.autoTable({
        startY: y,
        theme: 'grid',
        head: controlHead,
        body: controlBody,
        styles: { fontSize: 8, cellPadding: 2, valign: 'middle' },
        headStyles: { fillColor: [210, 220, 230], textColor: 20, halign: 'left' },
        margin: { left: margin, right: margin },
        columnStyles: {
          0: { cellWidth: 30 },
          1: { cellWidth: 45 },
          2: { cellWidth: 25 },
          3: { cellWidth: 30 },
          4: { cellWidth: 25 },
          5: { cellWidth: 'auto' }
        }
      });
      y = doc.lastAutoTable.finalY + 6;

      // DATOS DEL VUELO
      const empresa = qs('empresa').value || '-';
      const codigoVuelo = qs('codigoVuelo').value || '-';
      const matricula = qs('matricula').value || '-';
      const origen = qs('origen').value || '-';
      const destino = qs('destino').value || '-';
      const horaPartida = qs('horaPartida').value || '';
      const horaArribo = qs('horaArribo').value || '';
      const horaShow = horaPartida || horaArribo || '-';
      const conDemora = qs('conDemora').value || 'NO';
      const tipoVuelo = document.querySelector('input[name="tipoVuelo"]:checked')?.value || 'Salida';
      const posicion = qs('posicion').value || '-';

      const vueloHead = [[
        'Empresa','Código de Vuelo','Matrícula Aeronave',
        'Origen','Destino','Hora (Part/Arr)',
        'Con demora','Tipo de Vuelo','Posición Plataforma'
      ]];
      const vueloBody = [[
        empresa, codigoVuelo, matricula,
        origen, destino, horaShow,
        conDemora, tipoVuelo, posicion
      ]];

      doc.autoTable({
        startY: y,
        head: vueloHead,
        body: vueloBody,
        styles: { fontSize: 8, cellPadding: 2 },
        headStyles: { fillColor: [210, 220, 230], textColor: 20 },
        margin: { left: margin, right: margin },
        columnStyles: {
          0: { cellWidth: 30 },
          1: { cellWidth: 25 },
          2: { cellWidth: 30 },
          3: { cellWidth: 25 },
          4: { cellWidth: 25 },
          5: { cellWidth: 25 },
          6: { cellWidth: 18 },
          7: { cellWidth: 25 },
          8: { cellWidth: 20 }
        },
        theme: 'grid'
      });
      y = doc.lastAutoTable.finalY + 4;

      // 👉 Resaltamos explícitamente el tipo de vuelo
      doc.setFontSize(9);
      doc.setFont(undefined, 'bold');
      doc.text('Tipo de Vuelo:', margin, y);
      doc.setFont(undefined, 'normal');
      doc.text(tipoVuelo, margin + 28, y);
      y += 6;

      // PERSONAL DE APOYO TERRESTRE (título claro)
      doc.setFontSize(10);
      doc.setFont(undefined, 'bold');
      doc.text('PERSONAL DE APOYO TERRESTRE', margin, y);
      y += 4;
      doc.setFontSize(8);
      doc.setFont(undefined, 'normal');

      const headPT = [['Apellido y Nombre','DNI','Legajo','Función','Grupo']];
      const bodyPT = [];
      $all('input[name="pt_nombre[]"]').forEach((el, i) => {
        bodyPT.push([
          (el.value || '-').substring(0, 50),
          ($all('input[name="pt_dni[]"]')[i] || {}).value || '-',
          ($all('input[name="pt_legajo[]"]')[i] || {}).value || '-',
          ($all('select[name="pt_funcion[]"]')[i] || {}).value || '-',
          ($all('input[name="pt_grupo[]"]')[i] || {}).value || '-'
        ]);
      });
      if (!bodyPT.length) {
        for (let i = 0; i < 4; i++) bodyPT.push(['','','','','']);
      }

      doc.autoTable({
        startY: y,
        head: headPT,
        body: bodyPT,
        margin: { left: margin, right: margin },
        styles: { fontSize: 8, cellPadding: 2 },
        headStyles: { fillColor: [210, 220, 230], textColor: 20 },
        columnStyles: {
          0: { cellWidth: 65 },
          1: { cellWidth: 30 },
          2: { cellWidth: 28 },
          3: { cellWidth: 35 },
          4: { cellWidth: 'auto' }
        },
        theme: 'grid'
      });
      y = doc.lastAutoTable.finalY + 6;

      // PERSONAL DE SEGURIDAD (título claro)
      doc.setFontSize(10);
      doc.setFont(undefined, 'bold');
      doc.text('PERSONAL DE SEGURIDAD', margin, y);
      y += 4;
      doc.setFontSize(8);
      doc.setFont(undefined, 'normal');

      const headPS = [['Apellido y Nombre','DNI','Legajo','Empresa']];
      const bodyPS = [];
      $all('input[name="ps_nombre[]"]').forEach((el, i) => {
        bodyPS.push([
          (el.value || '-').substring(0, 60),
          ($all('input[name="ps_dni[]"]')[i] || {}).value || '-',
          ($all('input[name="ps_legajo[]"]')[i] || {}).value || '-',
          ($all('input[name="ps_empresa[]"]')[i] || {}).value || '-'
        ]);
      });
      if (!bodyPS.length) {
        for (let i = 0; i < 3; i++) bodyPS.push(['','','','']);
      }

      doc.autoTable({
        startY: y,
        head: headPS,
        body: bodyPS,
        margin: { left: margin, right: margin },
        styles: { fontSize: 8, cellPadding: 2 },
        headStyles: { fillColor: [210, 220, 230], textColor: 20 },
        columnStyles: {
          0: { cellWidth: 70 },
          1: { cellWidth: 35 },
          2: { cellWidth: 35 },
          3: { cellWidth: 'auto' }
        },
        theme: 'grid'
      });
      y = doc.lastAutoTable.finalY + 6;

      // VEHÍCULOS (título claro)
      doc.setFontSize(10);
      doc.setFont(undefined, 'bold');
      doc.text('VEHÍCULOS CONTROLADOS', margin, y);
      y += 4;
      doc.setFontSize(8);
      doc.setFont(undefined, 'normal');

      const headVH = [['Tipo de Vehículo','Empresa','Nº Interno','Operador','Observaciones']];
      const bodyVH = [];
      $all('input[name="vh_tipo[]"]').forEach((el, i) => {
        bodyVH.push([
          (el.value || '-').substring(0, 40),
          ($all('input[name="vh_empresa[]"]')[i] || {}).value || '-',
          ($all('input[name="vh_interno[]"]')[i] || {}).value || '-',
          ($all('input[name="vh_operador[]"]')[i] || {}).value || '-',
          ($all('input[name="vh_obs[]"]')[i] || {}).value || '-'
        ]);
      });
      if (!bodyVH.length) {
        for (let i = 0; i < 2; i++) bodyVH.push(['','','','','']);
      }

      doc.autoTable({
        startY: y,
        head: headVH,
        body: bodyVH,
        margin: { left: margin, right: margin },
        styles: { fontSize: 8, cellPadding: 2 },
        headStyles: { fillColor: [210, 220, 230], textColor: 20 },
        columnStyles: {
          0: { cellWidth: 40 },
          1: { cellWidth: 40 },
          2: { cellWidth: 25 },
          3: { cellWidth: 40 },
          4: { cellWidth: 'auto' }
        },
        theme: 'grid'
      });
      y = doc.lastAutoTable.finalY + 6;

      // NOVEDADES
      const novRows = [
        ['Equipajes/Cargas', qs('novEquipajes').value || ''],
        ['Inspección del Personal', qs('novInspeccion').value || ''],
        ['Equipajes Utilizados', qs('novEquipajesUtil').value || ''],
        ['Otras Novedades', qs('otrasNovedades').value || '']
      ];

      doc.autoTable({
        startY: y,
        head: [['NOVEDADES','Detalle']],
        body: novRows,
        margin: { left: margin, right: margin },
        styles: { fontSize: 8, cellPadding: 3 },
        headStyles: { fillColor: [200, 210, 220], textColor: 20, halign: 'left' },
        columnStyles: {
          0: { cellWidth: 50 },
          1: { cellWidth: pageWidth - margin * 2 - 60 }
        },
        theme: 'grid'
      });
      y = doc.lastAutoTable.finalY + 8;

      // RESPONSABLE + FIRMA
      const headResp = [['Datos del Responsable PSA','']];
      const bodyResp = [
        ['Firma', ''],
        ['Aclaración', qs('respNombre').value || '-'],
        ['N° Legajo', qs('respLegajo').value || '-'],
        ['DNI', qs('respDNI').value || '-']
      ];

      let firmaCell = null;

      doc.autoTable({
        startY: y,
        head: headResp,
        body: bodyResp,
        margin: { left: margin, right: margin },
        styles: { fontSize: 9, cellPadding: 3 },
        headStyles: { fillColor: [200, 210, 220], textColor: 20, halign: 'center' },
        columnStyles: {
          0: { cellWidth: 35, fontStyle: 'bold' },
          1: { cellWidth: pageWidth - margin * 2 - 40 }
        },
        theme: 'grid',
        didDrawCell: function (data) {
          if (data.section === 'body' && data.row.index === 0 && data.column.index === 1) {
            firmaCell = {
              x: data.cell.x,
              y: data.cell.y,
              w: data.cell.width,
              h: data.cell.height,
              page: doc.internal.getCurrentPageInfo().pageNumber
            };
          }
        }
      });

      if (firmaCell) {
        doc.setPage(firmaCell.page);
        const padding = 2;
        const maxAncho = Math.min(firmaCell.w - padding * 2, 60);
        const altoFirma = 20;
        const sigW = maxAncho;
        const sigH = altoFirma;
        const sigX = firmaCell.x + (firmaCell.w - sigW) / 2;
        const sigY = firmaCell.y + (firmaCell.h - sigH) / 2;

        if (firmaImagen) {
          try {
            doc.addImage(firmaImagen, 'PNG', sigX, sigY, sigW, sigH);
          } catch (e) {
            console.warn('No se pudo dibujar la firma en el PDF', e);
          }
        } else {
          doc.setDrawColor(0);
          doc.rect(sigX, sigY, sigW, sigH);
        }
      }

      // Nombre del archivo: sigue usando ISO porque es más cómodo para ordenar
      const fechaForName = (fechaIso || '').replace(/-/g, '') || 'NOFECHA';
      const safeCodigo = (qs('codigoVuelo').value || 'NOCODE').replace(/\s+/g, '');
      const filename = `PLANILLA_BODEGA_${fechaForName}_${safeCodigo}.pdf`;
      doc.save(filename);

      alert('✓ PDF generado exitosamente!');
    } catch (err) {
      console.error(err);
      alert('Error al generar PDF: ' + (err.message || err));
    }
  }

  // Eventos globales
  window.addEventListener('load', init);

  window.addEventListener('click', (ev) => {
    if (ev.target === signatureModal) cerrarModalFirma();
    if (ev.target === previewModal) cerrarVistaPrevia();
  });

  // Debug
  window.__planilla = {
    generarPDF,
    guardarBorrador,
    updatePreview
  };

})();
