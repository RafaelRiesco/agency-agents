function callClaude(prompt) {
  const apiKey = PropertiesService.getScriptProperties()
                   .getProperty('ANTHROPIC_API_KEY');
  
  const url = 'https://api.anthropic.com/v1/messages';
  
  const payload = {
    model: 'claude-sonnet-4-20250514',
    max_tokens: 1000,
    messages: [
      { role: 'user', content: prompt }
    ]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());
  
  if (data.content && data.content[0]) {
    return data.content[0].text;
  } else {
    return 'Error: ' + JSON.stringify(data);
  }
}
function testAgente() {
  const dataMock = `
    Campaña: RTGT_CBO_Cajas_Zapato
    Plataforma: Meta Ads
    ROAS: 5.92
    Spend semanal: 17.442 CLP
    Frecuencia: 6.17
    CTR: 1.30%
    CPC: 437 CLP
    Categoría: Cajas de zapato
  `;

  const prompt = `
    Eres un estratega de paid media para cadacosaensulugar.cl, 
    un ecommerce chileno de organización del hogar.
    
    Benchmarks del negocio:
    - ROAS objetivo: 5.0
    - ROAS mínimo aceptable: 1.7
    - Frecuencia máxima: 5.0
    - CTR promedio cuenta: 1.80%
    
    Analiza esta campaña y responde:
    1. ¿Está funcionando bien o hay problemas?
    2. ¿Qué acción recomiendas tomar esta semana?
    
    Datos de la campaña:
    ${dataMock}
    
    Sé directo y específico. Máximo 5 líneas.
  `;

  const resultado = callClaude(prompt);
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
                  .getSheetByName('Análisis');
  sheet.getRange('A1').setValue('Test Día 2 — ' + new Date());
  sheet.getRange('A2').setValue(resultado);
  
  Logger.log(resultado);
}
function readMetaData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
                  .getSheetByName('Data Meta');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // Índices de columnas relevantes
  const colNombre = headers.indexOf('Nombre de la campaña');
  const colEstado = headers.indexOf('Entrega de la campaña');
  const colGasto = headers.indexOf('Importe gastado (CLP)');
  const colROAS = headers.indexOf('ROAS (retorno del gasto publicitario) de compras');
  const colCTR = headers.indexOf('CTR (tasa de clics en el enlace)');
  const colCPC = headers.indexOf('CPC (Coste por clic en el enlace) (CLP)');
  const colFrecuencia = headers.indexOf('Frecuencia');
  const colAlcance = headers.indexOf('Alcance');

  let resumen = '';
  let campanasActivas = 0;

  for (let i = 1; i < data.length; i++) {
    const fila = data[i];
    const gasto = parseFloat(fila[colGasto]) || 0;
    
    // Solo procesar campañas con gasto real
    if (gasto === 0) continue;
    
    campanasActivas++;
    const nombre = fila[colNombre];
    const estado = fila[colEstado];
    const roas = parseFloat(fila[colROAS]) || 0;
    const ctr = parseFloat(fila[colCTR]) || 0;
    const cpc = parseFloat(fila[colCPC]) || 0;
    const frecuencia = parseFloat(fila[colFrecuencia]) || 0;
    const alcance = parseInt(fila[colAlcance]) || 0;

    resumen += `
Campaña: ${nombre}
Estado: ${estado}
Gasto (CLP): ${gasto.toLocaleString()}
ROAS: ${roas.toFixed(2)}
CTR: ${ctr.toFixed(2)}%
CPC (CLP): ${cpc.toFixed(0)}
Frecuencia: ${frecuencia.toFixed(2)}
Alcance: ${alcance.toLocaleString()}
---`;
  }

  Logger.log(`Campañas con gasto: ${campanasActivas}`);
  Logger.log(resumen);
  return resumen;
}
function analizarCampanasReales() {
  const dataCampanas = readMetaData();
  
  const prompt = `
    Eres un estratega de paid media para cadacosaensulugar.cl,
    un ecommerce chileno de organización del hogar con 60% de margen bruto.
    
    Benchmarks del negocio:
    - ROAS objetivo: 5.0
    - ROAS mínimo aceptable (alerta roja): 1.7
    - ROAS prospecting mínimo: 3.0
    - Frecuencia máxima: 5.0
    - CTR promedio cuenta: 1.80%
    - Ticket promedio: 63.629 CLP
    
    Analiza estas campañas y responde las siguientes preguntas:
    1. ¿Qué campañas deben pausarse o revisarse urgente?
    2. ¿Qué campañas tienen potencial de escalar presupuesto?
    3. ¿Cuál es la acción de mayor palanca esta semana?
    
    Datos de campañas:
    ${dataCampanas}
    
    Sé directo y específico. Usa los benchmarks para justificar cada recomendación.
  `;

  const resultado = callClaude(prompt);
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
                  .getSheetByName('Análisis');
  sheet.getRange('A1').setValue('Análisis — ' + new Date());
  sheet.getRange('A2').setValue(resultado);
  
  Logger.log(resultado);
}
function getMetaAdsData() {
  const token = PropertiesService.getScriptProperties()
                  .getProperty('META_ACCESS_TOKEN');
  
  // Obtener el Ad Account ID
  const urlAccount = `https://graph.facebook.com/v19.0/me/adaccounts?fields=id,name&access_token=${token}`;
  const responseAccount = UrlFetchApp.fetch(urlAccount, {muteHttpExceptions: true});
  const dataAccount = JSON.parse(responseAccount.getContentText());
  
  Logger.log('Cuentas disponibles: ' + JSON.stringify(dataAccount));
  
  if (!dataAccount.data || dataAccount.data.length === 0) {
    Logger.log('Error: No se encontraron cuentas de anuncios');
    return null;
  }
  
  const adAccountId = dataAccount.data[0].id;
  Logger.log('Usando cuenta: ' + adAccountId);
  
  // Obtener métricas de campañas
  const fechaHoy = new Date();
  const fecha30dias = new Date(fechaHoy - 30 * 24 * 60 * 60 * 1000);
  const fechaInicio = fecha30dias.toISOString().split('T')[0];
  const fechaFin = fechaHoy.toISOString().split('T')[0];
  
  const fields = 'campaign_name,spend,impressions,clicks,ctr,cpc,frequency,reach,actions,action_values';
  const timeRange = encodeURIComponent(JSON.stringify({since: fechaInicio, until: fechaFin}));
  const url = `https://graph.facebook.com/v19.0/${adAccountId}/insights?fields=${fields}&time_range=${timeRange}&level=campaign&access_token=${token}`;
  
  const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  const data = JSON.parse(response.getContentText());
  
  Logger.log('Respuesta API: ' + JSON.stringify(data));
  return data;
}
function procesarMetaAPI() {
  const data = getMetaAdsData();
  if (!data || !data.data) {
    Logger.log('Sin datos');
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
                  .getSheetByName('Data Meta');
  
  // Limpiar sheet
  sheet.clearContents();
    sheet.getRange('B:H').setNumberFormat('0.00');
    sheet.getRange('B:D').setNumberFormat('0');
  
  // Headers
  sheet.getRange(1, 1, 1, 8).setValues([[
    'Campaña', 'Gasto (CLP)', 'Impresiones', 'Clicks', 
    'CTR', 'CPC (CLP)', 'Frecuencia', 'ROAS'
  ]]);

  let fila = 2;
  
  data.data.forEach(campana => {
    // Calcular ROAS
    let valorConversion = 0;
    let compras = 0;
    
    if (campana.action_values) {
      const purchaseAction = campana.action_values.find(
        a => a.action_type === 'offsite_conversion.fb_pixel_purchase'
      );
      if (purchaseAction) valorConversion = parseFloat(purchaseAction.value);
    }

    if (campana.actions) {
      const purchaseCount = campana.actions.find(
        a => a.action_type === 'offsite_conversion.fb_pixel_purchase'
      );
      if (purchaseCount) compras = parseFloat(purchaseCount.value);
    }

    const gasto = parseFloat(campana.spend) || 0;
    const roas = gasto > 0 ? (valorConversion / gasto).toFixed(2) : 0;

    if (gasto === 0) return; // Skip campañas sin gasto

    sheet.getRange(fila, 1, 1, 8).setValues([[
      String(campana.campaign_name),
      Number(gasto),
      Number(campana.impressions),
      Number(campana.clicks),
      String((parseFloat(campana.ctr)).toFixed(2) + '%'),
      Number(parseFloat(campana.cpc).toFixed(0)),
      Number(parseFloat(campana.frequency).toFixed(2)),
      Number(roas)
    ]]);
    fila++;
  });

  Logger.log('Datos escritos en Data Meta: ' + (fila - 2) + ' campañas');
}
function getMetaAdsetData() {
  const token = PropertiesService.getScriptProperties()
                  .getProperty('META_ACCESS_TOKEN');
  
  const urlAccount = `https://graph.facebook.com/v19.0/me/adaccounts?fields=id&access_token=${token}`;
  const responseAccount = UrlFetchApp.fetch(urlAccount, {muteHttpExceptions: true});
  const dataAccount = JSON.parse(responseAccount.getContentText());
  const adAccountId = dataAccount.data[0].id;

  const fechaHoy = new Date();
  const fecha7dias = new Date(fechaHoy - 7 * 24 * 60 * 60 * 1000);
  const fechaInicio = fecha7dias.toISOString().split('T')[0];
  const fechaFin = fechaHoy.toISOString().split('T')[0];

  const fields = 'campaign_name,adset_name,spend,impressions,clicks,ctr,cpc,frequency,reach,actions,action_values';
  const timeRange = encodeURIComponent(JSON.stringify({since: fechaInicio, until: fechaFin}));
  const url = `https://graph.facebook.com/v19.0/${adAccountId}/insights?fields=${fields}&time_range=${timeRange}&level=adset&access_token=${token}`;

  const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  const data = JSON.parse(response.getContentText());

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
                  .getSheetByName('Data Meta');

  // Escribir desde fila 20 para no pisar los datos de campaña
  sheet.getRange('A20:I20').setValues([[
    'Campaña', 'Adset', 'Gasto (CLP)', 'Impresiones',
    'Clicks', 'CTR', 'CPC (CLP)', 'Frecuencia', 'ROAS'
  ]]);

  let fila = 21;

  data.data.forEach(adset => {
    const gasto = parseFloat(adset.spend) || 0;
    if (gasto === 0) return;

    let valorConversion = 0;
    if (adset.action_values) {
      const purchase = adset.action_values.find(
        a => a.action_type === 'offsite_conversion.fb_pixel_purchase'
      );
      if (purchase) valorConversion = parseFloat(purchase.value);
    }

    const roas = gasto > 0 ? (valorConversion / gasto).toFixed(2) : 0;

    sheet.getRange(fila, 1, 1, 9).setValues([[
      String(adset.campaign_name),
      String(adset.adset_name),
      Number(gasto),
      Number(adset.impressions),
      Number(adset.clicks),
      String(parseFloat(adset.ctr).toFixed(2) + '%'),
      Number(parseFloat(adset.cpc).toFixed(0)),
      Number(parseFloat(adset.frequency).toFixed(2)),
      Number(roas)
    ]]);
    fila++;
  });

  Logger.log('Adsets escritos: ' + (fila - 21));
}
function getMetaAdData() {
  const token = PropertiesService.getScriptProperties()
                  .getProperty('META_ACCESS_TOKEN');
  
  const urlAccount = `https://graph.facebook.com/v19.0/me/adaccounts?fields=id&access_token=${token}`;
  const responseAccount = UrlFetchApp.fetch(urlAccount, {muteHttpExceptions: true});
  const dataAccount = JSON.parse(responseAccount.getContentText());
  const adAccountId = dataAccount.data[0].id;

  const fechaHoy = new Date();
  const fecha7dias = new Date(fechaHoy - 7 * 24 * 60 * 60 * 1000);
  const fechaInicio = fecha7dias.toISOString().split('T')[0];
  const fechaFin = fechaHoy.toISOString().split('T')[0];

  const fields = 'campaign_name,adset_name,ad_name,spend,impressions,clicks,ctr,cpc,frequency,reach,actions,action_values';
  const timeRange = encodeURIComponent(JSON.stringify({since: fechaInicio, until: fechaFin}));
  const url = `https://graph.facebook.com/v19.0/${adAccountId}/insights?fields=${fields}&time_range=${timeRange}&level=ad&access_token=${token}`;

  const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  const data = JSON.parse(response.getContentText());

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
                  .getSheetByName('Data Meta');

  // Escribir desde fila 40 para no pisar campañas ni adsets
  sheet.getRange('A40:J40').setValues([[
    'Campaña', 'Adset', 'Anuncio', 'Gasto (CLP)', 'Impresiones',
    'Clicks', 'CTR', 'CPC (CLP)', 'Frecuencia', 'ROAS'
  ]]);

  let fila = 41;

  data.data.forEach(ad => {
    const gasto = parseFloat(ad.spend) || 0;
    if (gasto === 0) return;

    let valorConversion = 0;
    if (ad.action_values) {
      const purchase = ad.action_values.find(
        a => a.action_type === 'offsite_conversion.fb_pixel_purchase'
      );
      if (purchase) valorConversion = parseFloat(purchase.value);
    }

    const roas = gasto > 0 ? (valorConversion / gasto).toFixed(2) : 0;

    sheet.getRange(fila, 1, 1, 10).setValues([[
      String(ad.campaign_name),
      String(ad.adset_name),
      String(ad.ad_name),
      Number(gasto),
      Number(ad.impressions),
      Number(ad.clicks),
      String(parseFloat(ad.ctr).toFixed(2) + '%'),
      Number(parseFloat(ad.cpc).toFixed(0)),
      Number(parseFloat(ad.frequency).toFixed(2)),
      Number(roas)
    ]]);
    fila++;
  });

  Logger.log('Anuncios escritos: ' + (fila - 41));
}
function unificarData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
                  .getSheetByName('Data Meta');
  
  // Leer campañas (fila 1 headers, filas 2-15)
  const headersCampana = sheet.getRange(1, 1, 1, 8).getValues()[0];
  const datosCampana = sheet.getRange(2, 1, 15, 8).getValues()
    .filter(fila => fila[0] !== '');

  // Leer adsets (fila 20 headers, filas 21-35)
  const headersAdset = sheet.getRange(20, 1, 1, 9).getValues()[0];
  const datosAdset = sheet.getRange(21, 1, 15, 9).getValues()
    .filter(fila => fila[0] !== '');

  // Leer anuncios (fila 40 headers, filas 41-80)
  const headersAnuncio = sheet.getRange(40, 1, 1, 10).getValues()[0];
  const datosAnuncio = sheet.getRange(41, 1, 40, 10).getValues()
    .filter(fila => fila[0] !== '');

  // Formatear campañas
  let textoCampanas = '## NIVEL CAMPAÑA\n';
  datosCampana.forEach(fila => {
    textoCampanas += `
Campaña: ${fila[0]}
Gasto: ${fila[1].toLocaleString()} CLP
Impresiones: ${fila[2].toLocaleString()}
Clicks: ${fila[3].toLocaleString()}
CTR: ${fila[4]}
CPC: ${fila[5]} CLP
Frecuencia: ${fila[6]}
ROAS: ${fila[7]}
---`;
  });

  // Formatear adsets
  let textoAdsets = '\n## NIVEL ADSET\n';
  datosAdset.forEach(fila => {
    textoAdsets += `
Campaña: ${fila[0]} | Adset: ${fila[1]}
Gasto: ${fila[2].toLocaleString()} CLP | CTR: ${fila[5]} | CPC: ${fila[6]} CLP
Frecuencia: ${fila[7]} | ROAS: ${fila[8]}
---`;
  });

  // Formatear anuncios — solo los relevantes (gasto > 0 y ROAS conocido)
  let textoAnuncios = '\n## NIVEL ANUNCIO\n';
  datosAnuncio
    .filter(fila => fila[3] > 0 && fila[9] > 0)
    .sort((a, b) => b[9] - a[9]) // ordenar por ROAS descendente
    .forEach(fila => {
      textoAnuncios += `
Anuncio: ${fila[2]} | Adset: ${fila[1]}
Gasto: ${fila[3].toLocaleString()} CLP | CTR: ${fila[6]} | ROAS: ${fila[9]}
---`;
    });

  const textoCompleto = textoCampanas + textoAdsets + textoAnuncios;
  Logger.log(textoCompleto);
  return textoCompleto;
}
function runFullAnalysis() {
  const data = unificarData();
  
  const configSheet = SpreadsheetApp.getActiveSpreadsheet()
                        .getSheetByName('Config');
  const configData = configSheet.getDataRange().getValues();
  
  // Leer config
  let config = {};
  configData.forEach(fila => {
    if (fila[0] && fila[1]) config[fila[0]] = fila[1];
  });

  const prompt = `
Eres un estratega senior de paid media para ${config['Negocio'] || 'cadacosaensulugar.cl'}, un ecommerce chileno de organización del hogar con ${config['Margen bruto promedio'] || '60%'} de margen bruto.

BENCHMARKS DEL NEGOCIO:
- ROAS objetivo: ${config['ROAS objetivo retargeting'] || '5.0'}
- ROAS mínimo prospecting: ${config['ROAS objetivo prospecting'] || '3.0'}
- ROAS alerta roja: ${config['ROAS mínimo absoluto (alerta roja)'] || '1.7'}
- Frecuencia máxima: ${config['Frecuencia máxima retargeting'] || '5.0'}
- CTR promedio cuenta: ${config['CTR promedio Meta'] || '1.80%'}
- Ticket promedio: ${config['Ticket promedio (CLP)'] || '63629'} CLP

CONTEXTO ACTIVO: ${config['Contexto activo'] || 'Sin contexto especial'}

DATOS DE CAMPAÑAS (últimos 7 días):
${data}

Responde estas 5 preguntas con datos concretos:

1. ¿Qué campañas deben pausarse o revisarse urgente? (justifica con ROAS vs benchmark)
2. ¿Qué campañas tienen ROAS sobre objetivo y pueden escalar presupuesto?
3. ¿Qué categoría de producto está generando el mejor retorno?
4. ¿Qué anuncios específicos tienen mejor performance y deben replicarse?
5. ¿Cuál es la acción de mayor palanca esta semana?

Formato de respuesta:
- Usa los tres niveles (campaña, adset, anuncio) para justificar cada recomendación
- Sé directo y específico — nombra campañas y anuncios reales
- Máximo 15 líneas en total
`;

  const resultado = callClaude(prompt);
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
                  .getSheetByName('Análisis');
  
  // Limpiar análisis anterior
  sheet.getRange('A1:B10').clearContent();
  
  sheet.getRange('A1').setValue('Reporte semanal — ' + new Date().toLocaleDateString('es-CL'));
  sheet.getRange('A2').setValue(resultado);
  sheet.getRange('A2').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  
  Logger.log(resultado);
  
  SpreadsheetApp.getUi().alert('✅ Análisis completado. Revisa la pestaña Análisis.');
}
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🤖 Agente Marketing')
    .addItem('Correr análisis semanal', 'runFullAnalysis')
    .addItem('Actualizar data Meta', 'procesarMetaAPI')
    .addItem('Actualizar adsets', 'getMetaAdsetData')
    .addItem('Actualizar anuncios', 'getMetaAdData')
    .addToUi();
}