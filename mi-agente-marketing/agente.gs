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
  
  const urlAccount = `https://graph.facebook.com/v19.0/me/adaccounts?fields=id&access_token=${token}`;
  const responseAccount = UrlFetchApp.fetch(urlAccount, {muteHttpExceptions: true});
  const dataAccount = JSON.parse(responseAccount.getContentText());
  
  if (!dataAccount.data || dataAccount.data.length === 0) {
    Logger.log('Error: No se encontraron cuentas');
    return null;
  }
  
  const adAccountId = dataAccount.data[0].id;

  // 1. Traer estado y fecha inicio desde endpoint de campañas
  const urlCampanas = `https://graph.facebook.com/v19.0/${adAccountId}/campaigns?fields=name,configured_status,start_time&access_token=${token}`;
  const responseCampanas = UrlFetchApp.fetch(urlCampanas, {muteHttpExceptions: true});
  const dataCampanas = JSON.parse(responseCampanas.getContentText());
  
  // Crear mapa nombre → {estado, diasActiva}
  const mapaCampanas = {};
  if (dataCampanas.data) {
    dataCampanas.data.forEach(c => {
      const diasActiva = c.start_time 
        ? Math.floor((new Date() - new Date(c.start_time)) / (1000 * 60 * 60 * 24))
        : 0;
      mapaCampanas[c.name] = {
        estado: c.configured_status,
        diasActiva: diasActiva
      };
    });
  }

  // 2. Traer métricas de insights
  const fechaHoy = new Date();
  const fecha7dias = new Date(fechaHoy - 7 * 24 * 60 * 60 * 1000);
  const fechaInicio = fecha7dias.toISOString().split('T')[0];
  const fechaFin = fechaHoy.toISOString().split('T')[0];
  
  const fields = 'campaign_name,spend,impressions,clicks,ctr,cpc,frequency,reach,actions,action_values';
  const timeRange = encodeURIComponent(JSON.stringify({since: fechaInicio, until: fechaFin}));
  const url = `https://graph.facebook.com/v19.0/${adAccountId}/insights?fields=${fields}&time_range=${timeRange}&level=campaign&access_token=${token}`;
  
  const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  const data = JSON.parse(response.getContentText());

  // 3. Combinar insights con estado y días activa
  if (data.data) {
    data.data = data.data.map(c => {
      const info = mapaCampanas[c.campaign_name] || {estado: 'unknown', diasActiva: 0};
      return {...c, configured_status: info.estado, dias_activa: info.diasActiva};
    });
  }

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
  
  sheet.clearContents();
  
  sheet.getRange(1, 1, 1, 10).setValues([[
    'Campaña', 'Estado', 'Días activa', 'Gasto (CLP)', 'Impresiones', 
    'Clicks', 'CTR', 'CPC (CLP)', 'Frecuencia', 'ROAS'
  ]]);

  let fila = 2;
  
  data.data.forEach(campana => {
    const gasto = parseFloat(campana.spend) || 0;
    if (gasto === 0) return;

    const estado = campana.configured_status || 'unknown';
    const diasActiva = campana.dias_activa || 0;

    let valorConversion = 0;
    if (campana.action_values) {
      const purchase = campana.action_values.find(
        a => a.action_type === 'offsite_conversion.fb_pixel_purchase'
      );
      if (purchase) valorConversion = parseFloat(purchase.value);
    }

    const roas = gasto > 0 ? (valorConversion / gasto).toFixed(2) : 0;

    sheet.getRange(fila, 1, 1, 10).setValues([[
      String(campana.campaign_name),
      String(estado),
      Number(diasActiva),
      Number(gasto),
      Number(campana.impressions || 0),
      Number(campana.clicks || 0),
      String(parseFloat(campana.ctr || 0).toFixed(2) + '%'),
      Number(parseFloat(campana.cpc || 0).toFixed(0)),
      Number(parseFloat(campana.frequency || 0).toFixed(2)),
      Number(roas)
    ]]);
    fila++;
  });

  Logger.log('Campañas escritas: ' + (fila - 2));
  SpreadsheetApp.getUi().alert('✅ Data Meta actualizada: ' + (fila - 2) + ' campañas.');
}
function getMetaAdsetData() {
  const token = PropertiesService.getScriptProperties()
                  .getProperty('META_ACCESS_TOKEN');
  
  const urlAccount = `https://graph.facebook.com/v19.0/me/adaccounts?fields=id&access_token=${token}`;
  const responseAccount = UrlFetchApp.fetch(urlAccount, {muteHttpExceptions: true});
  const dataAccount = JSON.parse(responseAccount.getContentText());
  
  if (!dataAccount.data || dataAccount.data.length === 0) {
    SpreadsheetApp.getUi().alert('❌ Error de autenticación. Token expirado.');
    return;
  }
  
  const adAccountId = dataAccount.data[0].id;

  // 1. Traer estado y fecha inicio de adsets
  const urlAdsets = `https://graph.facebook.com/v19.0/${adAccountId}/adsets?fields=name,configured_status,start_time&access_token=${token}`;
  const responseAdsets = UrlFetchApp.fetch(urlAdsets, {muteHttpExceptions: true});
  const dataAdsets = JSON.parse(responseAdsets.getContentText());
  
  const mapaAdsets = {};
  if (dataAdsets.data) {
    dataAdsets.data.forEach(a => {
      const diasActiva = a.start_time
        ? Math.floor((new Date() - new Date(a.start_time)) / (1000 * 60 * 60 * 24))
        : 0;
      mapaAdsets[a.name] = {
        estado: a.configured_status,
        diasActiva: diasActiva
      };
    });
  }

  // 2. Traer métricas
  const fechaHoy = new Date();
  const fecha7dias = new Date(fechaHoy - 7 * 24 * 60 * 60 * 1000);
  const fechaInicio = fecha7dias.toISOString().split('T')[0];
  const fechaFin = fechaHoy.toISOString().split('T')[0];

  const fields = 'campaign_name,adset_name,spend,impressions,clicks,ctr,cpc,frequency,reach,actions,action_values';
  const timeRange = encodeURIComponent(JSON.stringify({since: fechaInicio, until: fechaFin}));
  const url = `https://graph.facebook.com/v19.0/${adAccountId}/insights?fields=${fields}&time_range=${timeRange}&level=adset&access_token=${token}`;

  const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  const data = JSON.parse(response.getContentText());

  if (!data.data) {
    SpreadsheetApp.getUi().alert('❌ Error API adsets: ' + JSON.stringify(data.error || data));
    return;
  }

  // 3. Combinar
  data.data = data.data.map(a => {
    const info = mapaAdsets[a.adset_name] || {estado: 'unknown', diasActiva: 0};
    return {...a, configured_status: info.estado, dias_activa: info.diasActiva};
  });

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
                  .getSheetByName('Data Meta');

  sheet.getRange('A20:K20').setValues([[
    'Campaña', 'Adset', 'Estado', 'Días activa', 'Gasto (CLP)', 
    'Impresiones', 'Clicks', 'CTR', 'CPC (CLP)', 'Frecuencia', 'ROAS'
  ]]);

  let fila = 21;

  data.data.forEach(adset => {
    const gasto = parseFloat(adset.spend) || 0;
    if (gasto === 0) return;

    const estado = adset.configured_status || 'unknown';
    const diasActiva = adset.dias_activa || 0;

    let valorConversion = 0;
    if (adset.action_values) {
      const purchase = adset.action_values.find(
        a => a.action_type === 'offsite_conversion.fb_pixel_purchase'
      );
      if (purchase) valorConversion = parseFloat(purchase.value);
    }

    const roas = gasto > 0 ? (valorConversion / gasto).toFixed(2) : 0;

    sheet.getRange(fila, 1, 1, 11).setValues([[
      String(adset.campaign_name),
      String(adset.adset_name),
      String(estado),
      Number(diasActiva),
      Number(gasto),
      Number(adset.impressions || 0),
      Number(adset.clicks || 0),
      String(parseFloat(adset.ctr || 0).toFixed(2) + '%'),
      Number(parseFloat(adset.cpc || 0).toFixed(0)),
      Number(parseFloat(adset.frequency || 0).toFixed(2)),
      Number(roas)
    ]]);
    fila++;
  });

  Logger.log('Adsets escritos: ' + (fila - 21));
  SpreadsheetApp.getUi().alert('✅ Adsets actualizados: ' + (fila - 21));
}
function getMetaAdData() {
  const token = PropertiesService.getScriptProperties()
                  .getProperty('META_ACCESS_TOKEN');
  
  const urlAccount = `https://graph.facebook.com/v19.0/me/adaccounts?fields=id&access_token=${token}`;
  const responseAccount = UrlFetchApp.fetch(urlAccount, {muteHttpExceptions: true});
  const dataAccount = JSON.parse(responseAccount.getContentText());
  
  if (!dataAccount.data || dataAccount.data.length === 0) {
    SpreadsheetApp.getUi().alert('❌ Error de autenticación. Token expirado.');
    return;
  }
  
  const adAccountId = dataAccount.data[0].id;

  // 1. Traer estado y fecha inicio de anuncios
  const urlAds = `https://graph.facebook.com/v19.0/${adAccountId}/ads?fields=name,configured_status,created_time&access_token=${token}`;
  const responseAds = UrlFetchApp.fetch(urlAds, {muteHttpExceptions: true});
  const dataAds = JSON.parse(responseAds.getContentText());
  
  const mapaAds = {};
  if (dataAds.data) {
    dataAds.data.forEach(a => {
      const diasActiva = a.created_time
        ? Math.floor((new Date() - new Date(a.created_time)) / (1000 * 60 * 60 * 24))
        : 0;
      mapaAds[a.name] = {
        estado: a.configured_status,
        diasActiva: diasActiva
      };
    });
  }

  // 2. Traer métricas
  const fechaHoy = new Date();
  const fecha7dias = new Date(fechaHoy - 7 * 24 * 60 * 60 * 1000);
  const fechaInicio = fecha7dias.toISOString().split('T')[0];
  const fechaFin = fechaHoy.toISOString().split('T')[0];

  const fields = 'campaign_name,adset_name,ad_name,spend,impressions,clicks,ctr,cpc,frequency,reach,actions,action_values';
  const timeRange = encodeURIComponent(JSON.stringify({since: fechaInicio, until: fechaFin}));
  const url = `https://graph.facebook.com/v19.0/${adAccountId}/insights?fields=${fields}&time_range=${timeRange}&level=ad&access_token=${token}`;

  const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  const data = JSON.parse(response.getContentText());

  if (!data.data) {
    SpreadsheetApp.getUi().alert('❌ Error API anuncios: ' + JSON.stringify(data.error || data));
    return;
  }

  // 3. Combinar
  data.data = data.data.map(a => {
    const info = mapaAds[a.ad_name] || {estado: 'unknown', diasActiva: 0};
    return {...a, configured_status: info.estado, dias_activa: info.diasActiva};
  });

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
                  .getSheetByName('Data Meta');

  sheet.getRange('A40:L40').setValues([[
    'Campaña', 'Adset', 'Anuncio', 'Estado', 'Días activa',
    'Gasto (CLP)', 'Impresiones', 'Clicks', 'CTR', 'CPC (CLP)', 'Frecuencia', 'ROAS'
  ]]);

  let fila = 41;

  data.data.forEach(ad => {
    const gasto = parseFloat(ad.spend) || 0;
    if (gasto === 0) return;

    const estado = ad.configured_status || 'unknown';
    const diasActiva = ad.dias_activa || 0;

    let valorConversion = 0;
    if (ad.action_values) {
      const purchase = ad.action_values.find(
        a => a.action_type === 'offsite_conversion.fb_pixel_purchase'
      );
      if (purchase) valorConversion = parseFloat(purchase.value);
    }

    const roas = gasto > 0 ? (valorConversion / gasto).toFixed(2) : 0;

    sheet.getRange(fila, 1, 1, 12).setValues([[
      String(ad.campaign_name),
      String(ad.adset_name),
      String(ad.ad_name),
      String(estado),
      Number(diasActiva),
      Number(gasto),
      Number(ad.impressions || 0),
      Number(ad.clicks || 0),
      String(parseFloat(ad.ctr || 0).toFixed(2) + '%'),
      Number(parseFloat(ad.cpc || 0).toFixed(0)),
      Number(parseFloat(ad.frequency || 0).toFixed(2)),
      Number(roas)
    ]]);
    fila++;
  });

  Logger.log('Anuncios escritos: ' + (fila - 41));
  SpreadsheetApp.getUi().alert('✅ Anuncios actualizados: ' + (fila - 41));
}
function unificarData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
                  .getSheetByName('Data Meta');

  // Leer contexto activo
  const configSheet = SpreadsheetApp.getActiveSpreadsheet()
                        .getSheetByName('Config');
  const configData = configSheet.getDataRange().getValues();
  let contextoActivo = 'Sin contexto especial';
  configData.forEach(fila => {
    if (fila[0] === 'Contexto activo') contextoActivo = fila[1];
  });

  // Campañas: 0=Nombre, 1=Estado, 2=Días, 3=Gasto, 4=Impresiones
  //           5=Clicks, 6=CTR, 7=CPC, 8=Frecuencia, 9=ROAS
  const datosCampana = sheet.getRange(2, 1, 15, 10).getValues()
    .filter(fila => fila[0] !== '');

  // Adsets: 0=Campaña, 1=Adset, 2=Estado, 3=Días, 4=Gasto
  //         5=Impresiones, 6=Clicks, 7=CTR, 8=CPC, 9=Frecuencia, 10=ROAS
  const datosAdset = sheet.getRange(21, 1, 15, 11).getValues()
    .filter(fila => fila[0] !== '');

  // Anuncios: 0=Campaña, 1=Adset, 2=Anuncio, 3=Estado, 4=Días
  //           5=Gasto, 6=Impresiones, 7=Clicks, 8=CTR, 9=CPC, 10=Frecuencia, 11=ROAS
  const datosAnuncio = sheet.getRange(41, 1, 40, 12).getValues()
    .filter(fila => fila[0] !== '');

  let texto = `## CONTEXTO ACTIVO\n${contextoActivo}\n\n`;

  // Campañas
  texto += '## NIVEL CAMPAÑA\n';
  datosCampana.forEach(fila => {
    const enAprendizaje = fila[2] < 7 ? '⚠️ EN APRENDIZAJE' : '';
    const estadoIcon = fila[1] === 'PAUSED' ? '⏸️ PAUSADA' : '▶️ ACTIVA';
    texto += `
Campaña: ${fila[0]}
Estado: ${estadoIcon} ${enAprendizaje} | Días activa: ${fila[2]}
Gasto: ${Number(fila[3]).toLocaleString()} CLP | CTR: ${fila[6]} | CPC: ${fila[7]} CLP | Frecuencia: ${fila[8]} | ROAS: ${fila[9]}
---`;
  });

  // Adsets
  texto += '\n## NIVEL ADSET\n';
  datosAdset.forEach(fila => {
    const enAprendizaje = fila[3] < 7 ? '⚠️ EN APRENDIZAJE' : '';
    const estadoIcon = fila[2] === 'PAUSED' ? '⏸️ PAUSADO' : '▶️ ACTIVO';
    texto += `
Campaña: ${fila[0]} | Adset: ${fila[1]}
Estado: ${estadoIcon} ${enAprendizaje} | Días activa: ${fila[3]}
Gasto: ${Number(fila[4]).toLocaleString()} CLP | CTR: ${fila[7]} | CPC: ${fila[8]} CLP | Frecuencia: ${fila[9]} | ROAS: ${fila[10]}
---`;
  });

  // Anuncios
  texto += '\n## NIVEL ANUNCIO\n';
  datosAnuncio
    .filter(fila => fila[5] > 0 && fila[11] > 0)
    .sort((a, b) => b[11] - a[11])
    .forEach(fila => {
      const enAprendizaje = fila[4] < 7 ? '⚠️ EN APRENDIZAJE' : '';
      const estadoIcon = fila[3] === 'PAUSED' ? '⏸️ PAUSADO' : '▶️ ACTIVO';
      texto += `
Anuncio: ${fila[2]} | Adset: ${fila[1]}
Estado: ${estadoIcon} ${enAprendizaje} | Días activa: ${fila[4]}
Gasto: ${Number(fila[5]).toLocaleString()} CLP | CTR: ${fila[8]} | ROAS: ${fila[11]}
---`;
    });

  Logger.log(texto);
  return texto;
}
function runFullAnalysis() {
  const data = unificarData();
  
  // Leer Config
  const configSheet = SpreadsheetApp.getActiveSpreadsheet()
                        .getSheetByName('Config');
  const configData = configSheet.getDataRange().getValues();
  let config = {};
  configData.forEach(fila => {
    if (fila[0] && fila[1]) config[fila[0]] = fila[1];
  });

  // Leer Memoria
  const memoriaSheet = SpreadsheetApp.getActiveSpreadsheet()
                         .getSheetByName('Memoria');
  const memoriaData = memoriaSheet.getDataRange().getValues();
  let memoria = '';
  memoriaData.forEach(fila => {
    if (fila[0] && fila[1]) memoria += `${fila[0]}: ${fila[1]}\n`;
  });

  const systemPrompt = `
Eres el estratega senior de paid media de ${config['Negocio'] || 'cadacosaensulugar.cl'}.

Tienes dos capas de conocimiento que debes aplicar simultáneamente:

---
## CAPA 1 — CONOCIMIENTO TÉCNICO DE META ADS

Eres un especialista full-funnel en Meta Ads con dominio profundo de:

ARQUITECTURA DE CAMPAÑAS:
- CBO (Campaign Budget Optimization): el presupuesto se asigna a nivel campaña y Meta lo distribuye automáticamente entre adsets. NUNCA recomendar mover presupuesto manualmente entre adsets dentro de un CBO — eso lo gestiona Meta. La palanca en CBO es pausar adsets que drenan presupuesto sin convertir, lo que libera budget para los adsets ganadores automáticamente.
- ABO (Ad Set Budget Optimization): presupuesto fijo por adset, útil para testing controlado donde necesitas garantizar que cada adset recibe inversión mínima.
- Advantage+: campañas automatizadas donde Meta controla audiencia, placement y optimización. Menos control pero potencialmente más eficiente a escala.

FASE DE APRENDIZAJE:
- Meta necesita ~50 eventos de conversión por adset para completar el aprendizaje.
- Durante aprendizaje (generalmente 7-14 días) el ROAS es inestable — no tomar decisiones de pausa basadas en ROAS bajo en este período.
- Cualquier cambio significativo (presupuesto >20%, audiencia, creativo principal) reinicia el aprendizaje. Evitar cambios frecuentes en campañas que están aprendiendo.
- Señal de alerta real durante aprendizaje: CPA más de 3x el objetivo o gasto sin ninguna conversión después de 5+ días.

FUNNEL Y AUDIENCIAS:
- TOFU (Top of Funnel): audiencias frías — intereses, lookalikes, broad. Objetivo: generar awareness y primeras visitas. ROAS esperado más bajo (1.5-3x). Excluir siempre compradores recientes y visitantes de los últimos 30 días.
- MOFU (Middle of Funnel): audiencias tibias — visitantes del sitio, viewers de video 50%+, engagers de página. ROAS esperado medio (2-4x).
- BOFU (Bottom of Funnel): audiencias calientes — visitantes recientes, add-to-cart sin compra, compradores para upsell. ROAS esperado alto (4x+). Frecuencia más alta aceptable (hasta 5x por semana).
- NUNCA comparar ROAS entre etapas distintas del funnel — son objetivos diferentes.

CREATIVOS Y FATIGA:
- Frecuencia sobre 5 en prospecting = señal de fatiga. Sobre 3 en TOFU = saturación temprana.
- CTR bajo con frecuencia alta = el creativo se agotó, no es problema de audiencia.
- CTR bajo con frecuencia baja = problema de relevancia del creativo o audiencia incorrecta.
- Un anuncio con ROAS muy alto pero menos de 50K CLP de gasto puede ser outlier estadístico — necesita más volumen antes de escalar masivamente.
- Thumb-stop rate objetivo: 25%+ de visualizaciones de 3 segundos.

ESCALADO:
- Para escalar un adset ganador dentro de CBO: aumentar presupuesto de la campaña máximo 20% cada 3-4 días para no reiniciar aprendizaje.
- Para escalar un anuncio ganador: duplicarlo en un nuevo adset con audiencia similar pero diferente segmento, no simplemente aumentar presupuesto del mismo adset.
- Señales para escalar: ROAS estable por 7+ días, frecuencia bajo 3, audiencia con reach disponible.

DIAGNÓSTICO DE PROBLEMAS:
- ROAS bajo + CTR alto + frecuencia baja = problema en la landing page o el proceso de compra, no en el anuncio.
- ROAS bajo + CTR bajo + frecuencia baja = problema de creativo o audiencia incorrecta.
- ROAS bajo + CTR alto + frecuencia alta = fatiga creativa, renovar anuncios.
- ROAS bajo + CTR bajo + frecuencia alta = audiencia agotada, expandir o cambiar segmento.

FECHAS ESPECIALES:
- Durante Black, Cyber y eventos de descuento el ROAS baja estructuralmente por el descuento en precio. ROAS mínimo aceptable en estos períodos: 2.5x.
- Campañas de eventos tienen fecha de término — no escalar agresivamente si quedan menos de 3 días.
- El valor real de una campaña de evento está en los datos de audiencia que genera, no solo en el ROAS inmediato.

---
## CAPA 2 — CONOCIMIENTO DEL NEGOCIO

${memoria}

---
## BENCHMARKS OPERATIVOS

- ROAS objetivo retargeting/BOFU: ${config['ROAS objetivo retargeting'] || '5.0'}
- ROAS objetivo prospecting/TOFU: ${config['ROAS objetivo prospecting'] || '3.0'}
- ROAS alerta roja (períodos normales): ${config['ROAS mínimo absoluto (alerta roja)'] || '1.7'}
- ROAS mínimo en Black/Cyber: 2.5
- Frecuencia máxima prospecting: 2.5
- Frecuencia máxima retargeting: ${config['Frecuencia máxima retargeting'] || '5.0'}
- CTR promedio cuenta: ${config['CTR promedio Meta'] || '1.80%'}
- Ticket promedio: ${config['Ticket promedio (CLP)'] || '63629'} CLP
- Margen bruto: ${config['Margen bruto promedio'] || '60%'}

---
## CONTEXTO ACTIVO

${config['Contexto activo'] || 'Sin contexto especial'}

---
## REGLAS DE ANÁLISIS

1. Antes de cualquier recomendación, identifica en qué etapa del funnel está cada campaña/adset y aplica los benchmarks correspondientes a esa etapa — no los generales.
2. Para campañas CBO, las recomendaciones de presupuesto van a nivel de campaña, nunca a nivel de adset individual.
3. Si una campaña o adset tiene menos de 7 días activo, márcalo como EN APRENDIZAJE y solo recomienda acción si hay señales de alarma extremas (0 conversiones con gasto significativo).
4. Antes de recomendar pausar un anuncio, verifica si tiene suficiente gasto para ser estadísticamente válido (mínimo 30K CLP).
5. Cruza siempre los tres niveles antes de concluir: campaña → adset → anuncio.
6. Ignora campos de Memoria marcados como POR EXPLORAR en el análisis.
`;

  const userPrompt = `
DATOS DE CAMPAÑAS (últimos 7 días):
${data}

Analiza en profundidad y responde estas 5 preguntas:

1. DIAGNÓSTICO POR CAMPAÑA: Para cada campaña activa, diagnostica qué está pasando realmente — identifica la etapa del funnel, si está en aprendizaje, y el problema específico si lo hay (creativo, audiencia, landing, presupuesto). Para campañas pausadas con gasto reciente, evalúa si vale la pena reactivar.

2. ACCIONES EN CBO: Considerando que las campañas activas son CBO, ¿qué adsets deberían pausarse para liberar presupuesto hacia los ganadores? Justifica con datos.

3. CREATIVOS — DIAGNÓSTICO Y ACCIÓN: ¿Qué anuncios están funcionando y por qué (según el framework CTR/frecuencia/ROAS)? ¿Qué anuncios tienen señales de fatiga o problema? ¿Qué tipo de creativo nuevo recomendarías basado en los patrones ganadores?

4. AUDIENCIAS Y FUNNEL: ¿Hay gaps en el funnel actual? ¿Alguna etapa sin cobertura o con audiencia agotada? ¿Qué audiencia nueva vale la pena probar?

5. PRIORIDAD DE LA SEMANA: Una sola acción concreta y ejecutable hoy, con justificación basada en datos y contexto del negocio.

Sé directo y técnico. Usa los frameworks de diagnóstico. Máximo 25 líneas.
`;

  const payload = {
    model: 'claude-sonnet-4-20250514',
    max_tokens: 2000,
    system: systemPrompt,
    messages: [
      { role: 'user', content: userPrompt }
    ]
  };

  const apiKey = PropertiesService.getScriptProperties()
                   .getProperty('ANTHROPIC_API_KEY');

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

  const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', options);
  const result = JSON.parse(response.getContentText());
  const resultado = result.content[0].text;

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
                  .getSheetByName('Análisis');
  sheet.getRange('A1:B30').clearContent();
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