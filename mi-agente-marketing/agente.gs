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