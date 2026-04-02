const SPREADSHEET_ID = '1wC3o_yhzHg-xGmK-TLqm3grErQsQ0vOptjtl6pZj-sg';
const SHEET_NAME = 'Form Responses 1';
let cache = CacheService.getScriptCache();
const CACHE_DURATION = 300;

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('CSTH Dengue Research Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getDashboardData() {
  const cached = cache.get('dashboardData');
  if (cached) {
    return JSON.parse(cached);
  }
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const rawData = sheet.getDataRange().getValues();
  const headers = rawData[0].map(h => h.toString());
  const data = rawData.slice(1).filter(row => row.some(cell => cell !== ''));
  
  const idx = getColumnIndices(headers);
  const result = {
    totalPatients: data.length,
    lastUpdated: new Date().toISOString(),
    kpis: calculateKPIs(data, headers, idx),
    severityDistribution: calculateSeverity(data, idx),
    demographics: calculateDemographics(data, idx),
    symptomTrends: calculateSymptomTrends(data, idx),
    riskAssessment: calculateRiskAssessment(data, idx),
    dataQuality: calculateDataQuality(data, headers),
    timelineProgression: calculateTimelineProgression(data, idx),
    highRiskPatients: identifyHighRiskPatients(data, headers, idx),
    clinicalInsights: generateClinicalInsights(data, idx),
    comorbidityAnalysis: analyzeComorbidities(data, idx),
    medicationAnalysis: analyzeMedications(data, idx),
    cohortComparison: calculateCohortComparison(data, idx),
    cohortAnalysis: calculateCohortAnalysis(data, headers, idx),
    patientData: calculatePatientData(data, headers, idx),
    medications: analyzeMedications(data, idx),
    collectionProgress: calculateCollectionProgress(data, headers),
    riskAnalysis: calculateRiskAnalysis(data, headers, idx),
    warningSigns: calculateWarningSigns(data, headers, idx),
    criticalPhase: calculateCriticalPhase(data, idx),
    symptomPatterns: calculateSymptomPatterns(data, headers, idx),
    lifestyleImpact: calculateLifestyleImpact(data, headers, idx),
    labValues: calculateLabValues(data, idx),
    serology: calculateSerology(data, idx),
    outcomes: calculateOutcomes(data, idx),
    collectionMonitor: calculateCollectionMonitor(data, headers, idx)
  };
  
  cache.put('dashboardData', JSON.stringify(result), CACHE_DURATION);
  return result;
}

function getKPIsOnly() {
  const cached = cache.get('kpisOnly');
  if (cached) return JSON.parse(cached);
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const rawData = sheet.getDataRange().getValues();
  const headers = rawData[0].map(h => h.toString());
  const data = rawData.slice(1).filter(row => row.some(cell => cell !== ''));
  const idx = getColumnIndices(headers);
  
  const result = calculateKPIs(data, headers, idx);
  cache.put('kpisOnly', JSON.stringify(result), CACHE_DURATION);
  return result;
}

function getHighRiskPatients() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const rawData = sheet.getDataRange().getValues();
  const headers = rawData[0].map(h => h.toString());
  const data = rawData.slice(1).filter(row => row.some(cell => cell !== ''));
  const idx = getColumnIndices(headers);
  
  return identifyHighRiskPatients(data, headers, idx);
}

function getDataQuality() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const rawData = sheet.getDataRange().getValues();
  const headers = rawData[0].map(h => h.toString());
  const data = rawData.slice(1).filter(row => row.some(cell => cell !== ''));
  
  return calculateDataQuality(data, headers);
}

function getPatientList(page, pageSize) {
  page = page || 0;
  pageSize = pageSize || 20;
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const rawData = sheet.getDataRange().getValues();
  const headers = rawData[0].map(h => h.toString());
  const data = rawData.slice(1).filter(row => row.some(cell => cell !== ''));
  const idx = getColumnIndices(headers);
  
  const start = page * pageSize;
  const end = Math.min(start + pageSize, data.length);
  const pageData = data.slice(start, end);
  
  return {
    patients: pageData.map(function(row, i) {
      return {
        index: start + i,
        patientId: row[idx.patientId] || ('Patient ' + (start + i + 1)),
        age: parseFloat(row[idx.age]) || 0,
        gender: row[idx.gender] || '',
        severity: assessSeverity(row, idx),
        dayAdmitted: row[idx.dayOfIllness] || '',
        hasBleeding: row[idx.bleeding] === 'Yes',
        hasComorbidity: row[idx.comorbidities_ht] === 'Yes' || row[idx.comorbidities_dm] === 'Yes'
      };
    }),
    total: data.length,
    page: page,
    pageSize: pageSize,
    totalPages: Math.ceil(data.length / pageSize)
  };
}

function getDetailedPatient(patientIndex) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const rawData = sheet.getDataRange().getValues();
  const headers = rawData[0].map(h => h.toString());
  const data = rawData.slice(1).filter(row => row.some(cell => cell !== ''));
  const idx = getColumnIndices(headers);
  
  if (patientIndex >= 0 && patientIndex < data.length) {
    const row = data[patientIndex];
    const patient = {};
    headers.forEach(function(h, i) {
      patient[h] = row[i];
    });
    patient._index = patientIndex;
    patient._severity = assessSeverity(row, idx);
    return patient;
  }
  return null;
}

function calculateKPIs(data, headers, idx) {
  let completedForms = 0;
  let severeCases = 0, mildCases = 0, moderateCases = 0;
  let averageAge = 0, averageBMI = 0;
  let maleCount = 0, femaleCount = 0;
  let withComorbidities = 0, withBleeding = 0;
  let smoking = 0, alcohol = 0;
  let hypertension = 0, diabetes = 0;
  let nsaids = 0, antiplatelet = 0, steroids = 0;
  let totalDays = 0;
  
  data.forEach(function(row) {
    const age = parseFloat(row[idx.age]) || 0;
    const bmi = parseFloat(row[idx.bmi]) || 0;
    const gender = (row[idx.gender] || '').toLowerCase();
    const day = parseFloat(row[idx.dayOfIllness]) || 0;
    
    averageAge += age;
    averageBMI += bmi;
    totalDays += day;
    
    if (gender === 'male') maleCount++;
    else if (gender === 'female') femaleCount++;
    
    if (row[idx.comorbidities_ht] === 'Yes') { withComorbidities++; hypertension++; }
    if (row[idx.comorbidities_dm] === 'Yes') { withComorbidities++; diabetes++; }
    if (row[idx.comorbidities_ihd] === 'Yes') withComorbidities++;
    
    if (row[idx.bleeding] === 'Yes') withBleeding++;
    if ((row[idx.smoking] || '').toLowerCase().includes('current')) smoking++;
    if (row[idx.alcohol] === 'Yes') alcohol++;
    if (row[idx.nsaids] === 'Yes') nsaids++;
    if (row[idx.antiplatelet] === 'Yes') antiplatelet++;
    if (row[idx.steroids] === 'Yes') steroids++;
    
    const severity = assessSeverity(row, idx);
    if (severity === 'Severe') severeCases++;
    else if (severity === 'Moderate') moderateCases++;
    else mildCases++;
    
    completedForms++;
  });
  
  const n = data.length || 1;
  return {
    totalPatients: data.length,
    completedForms: completedForms,
    completionRate: ((completedForms / data.length) * 100).toFixed(1),
    severeCases: severeCases,
    moderateCases: moderateCases,
    mildCases: mildCases,
    averageAge: (averageAge / n).toFixed(1),
    averageBMI: (averageBMI / n).toFixed(1),
    averageDayAdmitted: (totalDays / n).toFixed(1),
    maleCount: maleCount,
    femaleCount: femaleCount,
    malePercent: ((maleCount / n) * 100).toFixed(1),
    femalePercent: ((femaleCount / n) * 100).toFixed(1),
    withComorbidities: withComorbidities,
    comorbidityPercent: ((withComorbidities / n) * 100).toFixed(1),
    withBleeding: withBleeding,
    bleedingPercent: ((withBleeding / n) * 100).toFixed(1),
    smoking: smoking,
    alcohol: alcohol,
    hypertension: hypertension,
    diabetes: diabetes,
    nsaids: nsaids,
    antiplatelet: antiplatelet,
    steroids: steroids,
    mortalityRisk: {
      highRisk: withComorbidities + withBleeding + data.filter(function(r) { return (parseFloat(r[idx.age]) || 0) > 60; }).length,
      percentage: (((withComorbidities + withBleeding) / n * 100)).toFixed(1)
    }
  };
}

function getColumnIndices(headers) {
  return {
    timestamp: headers.findIndex(function(h) { return h.includes('Timestamp'); }),
    patientId: headers.findIndex(function(h) { return h.includes('Patient ID'); }),
    phone: headers.findIndex(function(h) { return h.includes('Phone'); }),
    admissionDate: headers.findIndex(function(h) { return h.includes('admission'); }),
    dischargeDate: headers.findIndex(function(h) { return h.includes('discharge'); }),
    age: headers.findIndex(function(h) { return h.includes('Age'); }),
    gender: headers.findIndex(function(h) { return h.includes('Gender'); }),
    employment: headers.findIndex(function(h) { return h.includes('Employment'); }),
    education: headers.findIndex(function(h) { return h.includes('Educational'); }),
    bmi: headers.findIndex(function(h) { return h.includes('BMI'); }),
    priorDengue: headers.findIndex(function(h) { return h.includes('prior dengue'); }),
    comorbidities_ht: headers.findIndex(function(h) { return h.includes('Hypertension'); }),
    comorbidities_dm: headers.findIndex(function(h) { return h.includes('Diabetes'); }),
    comorbidities_ihd: headers.findIndex(function(h) { return h.includes('Ischaemic'); }),
    comorbidities_asthma: headers.findIndex(function(h) { return h.includes('Asthma'); }),
    comorbidities_ckd: headers.findIndex(function(h) { return h.includes('kidney'); }),
    comorbidities_cld: headers.findIndex(function(h) { return h.includes('liver'); }),
    comorbidities_immuno: headers.findIndex(function(h) { return h.includes('Immunosuppressive'); }),
    comorbidities_fatty: headers.findIndex(function(h) { return h.includes('fatty liver'); }),
    smoking: headers.findIndex(function(h) { return h.includes('Smoking'); }),
    alcohol: headers.findIndex(function(h) { return h.includes('Alcohol'); }),
    fever: headers.findIndex(function(h) { return h.toLowerCase().includes('fever') && !h.includes('14.'); }),
    headache: headers.findIndex(function(h) { return h.includes('Headache'); }),
    retroOrbital: headers.findIndex(function(h) { return h.includes('retro-orbital'); }),
    arthralgia: headers.findIndex(function(h) { return h.includes('Arthralgia'); }),
    myalgia: headers.findIndex(function(h) { return h.includes('Myalgia'); }),
    rash: headers.findIndex(function(h) { return h.includes('Rash'); }),
    nausea: headers.findIndex(function(h) { return h.includes('Nausea'); }),
    vomiting: headers.findIndex(function(h) { return h.includes('Vomiting'); }),
    abdominalPain: headers.findIndex(function(h) { return h.includes('abdominal pain'); }),
    diarrhea: headers.findIndex(function(h) { return h.includes('Diarrhoea'); }),
    cough: headers.findIndex(function(h) { return h.includes('Cough'); }),
    bleeding: headers.findIndex(function(h) { return h.includes('bleeding manifestations'); }),
    dayOfIllness: headers.findIndex(function(h) { return h.includes('day of illness') && h.includes('admitted'); }),
    symptomOnset: headers.findIndex(function(h) { return h.includes('symptom onset'); }),
    nsaids: headers.findIndex(function(h) { return h.includes('NSAIDs'); }),
    antiplatelet: headers.findIndex(function(h) { return h.includes('Anti-platelet'); }),
    steroids: headers.findIndex(function(h) { return h.includes('Steroids'); }),

    // ── Warning signs (section 15 & 17) ──────────────────────────
    lethargy: headers.findIndex(function(h) { return h.includes('Lethargy'); }),
    persistentVomiting: headers.findIndex(function(h) { return h.includes('15.3'); }),
    hctRiseFlag: headers.findIndex(function(h) { return h.includes('Rapid rise in PCV'); }),
    fluidAccumulation: headers.findIndex(function(h) { return h.includes('Clinically detected fluid accumulation') && !h.includes('what'); }),
    mucosalBleeding: headers.findIndex(function(h) { return h.includes('17.2. Mucosal bleeding'); }),

    // ── Platelets (section 19.1) ──────────────────────────────────
    plateletAdm: headers.findIndex(function(h) { return h.includes('19.1.1'); }),
    plateletLow: headers.findIndex(function(h) { return h.includes('19.1.2.') && !h.includes('19.1.2.1'); }),
    plateletDis: headers.findIndex(function(h) { return h.includes('19.1.3'); }),

    // ── HCT / PCV (section 19.2) ──────────────────────────────────
    pcvAdm: headers.findIndex(function(h) { return h.includes('On admission - PCV'); }),
    hctAdm: headers.findIndex(function(h) { return h.includes('On admission - HCT'); }),

    // ── WBC (section 19.3) ───────────────────────────────────────
    wbcAdm: headers.findIndex(function(h) { return h.includes('19.3.1'); }),
    wbcLow: headers.findIndex(function(h) { return h.includes('19.3.2.') && !h.includes('19.3.2.1'); }),

    // ── Liver enzymes: ALT (19.4), AST (19.5) ────────────────────
    altAdm:  headers.findIndex(function(h) { return h.includes('19.4.1'); }),
    altHigh: headers.findIndex(function(h) { return h.includes('19.4.2.') && !h.includes('19.4.2.1'); }),
    astAdm:  headers.findIndex(function(h) { return h.includes('19.5.1'); }),
    astHigh: headers.findIndex(function(h) { return h.includes('19.5.2.') && !h.includes('19.5.2.1'); }),

    // ── Renal / coagulation (19.6 – 19.8) ────────────────────────
    creatinine: headers.findIndex(function(h) { return h.includes('19.6.1'); }),
    albumin:    headers.findIndex(function(h) { return h.includes('Serum albumin'); }),
    pt:         headers.findIndex(function(h) { return h.includes('Prothrombin time'); }),
    inr:        headers.findIndex(function(h) { return h.includes('INR'); }),

    // ── Serology ─────────────────────────────────────────────────
    ns1: headers.findIndex(function(h) { return h.includes('Dengue NS1 Antigen') && !h.includes('day'); }),
    igg: headers.findIndex(function(h) { return h.includes('22.1. IgG'); }),
    igm: headers.findIndex(function(h) { return h.includes('22.2. IgM'); }),

    // ── Vital signs ───────────────────────────────────────────────
    tempAdm: headers.findIndex(function(h) { return h.includes('16.1.1'); }),
    sbpAdm:  headers.findIndex(function(h) { return h.includes('16.2.1.1'); }),

    // ── Outcomes ─────────────────────────────────────────────────
    icuAdmit:     headers.findIndex(function(h) { return h.includes('Need for ICU admission') && !h.includes('day') && !h.includes('duration'); }),
    hduAdmit:     headers.findIndex(function(h) { return h.includes('Need for HDU admission') && !h.includes('day') && !h.includes('duration'); }),
    hospitalDays: headers.findIndex(function(h) { return h.includes('Total duration of hospital stay'); }),
    finalDiagnosis: headers.findIndex(function(h) { return h.includes('Final diagnosis'); }),
    outcome:      headers.findIndex(function(h) { return h.includes('Outcome at discharge'); }),

    // ── Complications (section 30) ────────────────────────────────
    compBleeding:    headers.findIndex(function(h) { return h.includes('Severe bleeding'); }),
    compMOD:         headers.findIndex(function(h) { return h.includes('Multi-organ dysfunction'); }),
    compARDS:        headers.findIndex(function(h) { return h.includes('ARDS'); }),
    compALF:         headers.findIndex(function(h) { return h.includes('Acute Liver Failure'); }),
    compAKI:         headers.findIndex(function(h) { return h.includes('Acute Kidney Injury'); }),
    compMyocarditis: headers.findIndex(function(h) { return h.includes('Myocarditis'); }),
    compFluidOverload: headers.findIndex(function(h) { return h.includes('Fluid overload'); })
  };
}

function assessSeverity(row, idx) {
  const hasBleeding = row[idx.bleeding] === 'Yes';
  const dayAdmitted = parseFloat(row[idx.dayOfIllness]) || 0;
  const hasComorbidity = row[idx.comorbidities_ht] === 'Yes' || 
                         row[idx.comorbidities_dm] === 'Yes' ||
                         row[idx.comorbidities_ihd] === 'Yes';
  const age = parseFloat(row[idx.age]) || 0;
  
  if (hasBleeding || (hasComorbidity && dayAdmitted > 3) || age > 60 || age < 5) return 'Severe';
  if (dayAdmitted > 5 || hasComorbidity) return 'Moderate';
  return 'Mild';
}

function calculateSeverity(data, idx) {
  let severe = 0, moderate = 0, mild = 0;
  data.forEach(function(row) {
    const severity = assessSeverity(row, idx);
    if (severity === 'Severe') severe++;
    else if (severity === 'Moderate') moderate++;
    else mild++;
  });
  
  const total = data.length || 1;
  return {
    labels: ['Mild', 'Moderate', 'Severe'],
    values: [
      ((mild / total) * 100).toFixed(1),
      ((moderate / total) * 100).toFixed(1),
      ((severe / total) * 100).toFixed(1)
    ],
    counts: [mild, moderate, severe]
  };
}

function calculateDemographics(data, idx) {
  const ageGroups = { '0-10': 0, '11-20': 0, '21-30': 0, '31-40': 0, '41-50': 0, '51-60': 0, '60+': 0 };
  const bmiCategories = { 'Underweight': 0, 'Normal': 0, 'Overweight': 0, 'Obese': 0 };
  
  data.forEach(function(row) {
    const age = parseFloat(row[idx.age]) || 0;
    const bmi = parseFloat(row[idx.bmi]) || 0;
    
    if (age <= 10) ageGroups['0-10']++;
    else if (age <= 20) ageGroups['11-20']++;
    else if (age <= 30) ageGroups['21-30']++;
    else if (age <= 40) ageGroups['31-40']++;
    else if (age <= 50) ageGroups['41-50']++;
    else if (age <= 60) ageGroups['51-60']++;
    else ageGroups['60+']++;
    
    if (bmi < 18.5) bmiCategories['Underweight']++;
    else if (bmi < 25) bmiCategories['Normal']++;
    else if (bmi < 30) bmiCategories['Overweight']++;
    else bmiCategories['Obese']++;
  });
  
  return {
    ageGroups: ageGroups,
    bmiCategories: bmiCategories,
    genderDistribution: {
      male: data.filter(function(r) { return (r[idx.gender] || '').toLowerCase() === 'male'; }).length,
      female: data.filter(function(r) { return (r[idx.gender] || '').toLowerCase() === 'female'; }).length
    }
  };
}

function calculateSymptomTrends(data, idx) {
  const symptoms = ['fever', 'headache', 'retroOrbital', 'arthralgia', 'myalgia', 'rash', 'nausea', 'vomiting', 'abdominalPain', 'diarrhea', 'cough'];
  const symptomTrends = {};
  
  symptoms.forEach(function(symptom) {
    const colIdx = idx[symptom];
    if (colIdx >= 0) {
      const yesCount = data.filter(function(r) { return r[colIdx] === 'Yes'; }).length;
      symptomTrends[symptom] = {
        count: yesCount,
        percentage: ((yesCount / data.length) * 100).toFixed(1)
      };
    }
  });
  
  return symptomTrends;
}

function calculateRiskAssessment(data, idx) {
  const riskScores = [];
  
  data.forEach(function(row, i) {
    let score = 0;
    const factors = [];
    
    const age = parseFloat(row[idx.age]) || 0;
    if (age > 60 || age < 5) { score += 3; factors.push('Age risk'); }
    else if (age > 45) { score += 2; factors.push('Mid-age'); }
    
    if (row[idx.comorbidities_dm] === 'Yes') { score += 3; factors.push('Diabetes'); }
    if (row[idx.comorbidities_ht] === 'Yes') { score += 2; factors.push('Hypertension'); }
    if (row[idx.comorbidities_ihd] === 'Yes') { score += 2; factors.push('IHD'); }
    
    if (row[idx.bleeding] === 'Yes') { score += 3; factors.push('Bleeding'); }
    
    const day = parseFloat(row[idx.dayOfIllness]) || 0;
    if (day > 5) { score += 2; factors.push('Late admission'); }
    
    if (row[idx.nsaids] === 'Yes') { score += 2; factors.push('NSAID use'); }
    if (row[idx.antiplatelet] === 'Yes') { score += 2; factors.push('Antiplatelet'); }
    
    let risk = 'Low';
    if (score >= 8) risk = 'Critical';
    else if (score >= 5) risk = 'High';
    else if (score >= 3) risk = 'Moderate';
    
    riskScores.push({
      patientId: row[idx.patientId] || ('Patient ' + (i + 1)),
      index: i,
      score: score,
      risk: risk,
      factors: factors,
      age: age,
      hasBleeding: row[idx.bleeding] === 'Yes'
    });
  });
  
  return riskScores.sort(function(a, b) { return b.score - a.score; });
}

function calculateDataQuality(data, headers) {
  const quality = {};
  const totalRows = data.length;
  const keyFields = ['Age', 'Gender', 'BMI', 'Hypertension', 'Diabetes', 'Bleeding', 'Admission'];
  
  headers.forEach(function(header, i) {
    if (keyFields.some(function(k) { return header.includes(k); })) {
      const filledCount = data.filter(function(row) { return row[i] !== '' && row[i] !== null; }).length;
      const completeness = ((filledCount / totalRows) * 100).toFixed(1);
      
      quality[header] = {
        completeness: completeness,
        missing: totalRows - filledCount,
        status: completeness > 80 ? 'Good' : completeness > 50 ? 'Fair' : 'Poor'
      };
    }
  });
  
  return quality;
}

function calculateTimelineProgression(data, idx) {
  const daySymptoms = {};
  
  data.forEach(function(row) {
    const day = parseInt(row[idx.dayOfIllness]) || 0;
    if (!daySymptoms[day]) {
      daySymptoms[day] = { patients: 0, severity: { mild: 0, moderate: 0, severe: 0 } };
    }
    daySymptoms[day].patients++;
    
    const severity = assessSeverity(row, idx);
    daySymptoms[day].severity[severity.toLowerCase()]++;
  });
  
  return daySymptoms;
}

function identifyHighRiskPatients(data, headers, idx) {
  const highRisk = [];
  const pregIdx = headers.findIndex(function(h) { return h.toLowerCase().includes('pregnancy'); });
  const ckdIdx = headers.findIndex(function(h) { return h.toLowerCase().includes('chronic kidney'); });
  const immIdx = headers.findIndex(function(h) { return h.toLowerCase().includes('immunosuppressive'); });
  
  data.forEach(function(row, i) {
    const age = parseFloat(row[idx.age]) || 0;
    const hasBleeding = row[idx.bleeding] === 'Yes';
    const hasComorbidity = row[idx.comorbidities_ht] === 'Yes' || row[idx.comorbidities_dm] === 'Yes';
    const lateAdmission = parseFloat(row[idx.dayOfIllness]) > 5;
    
    const hasDM = row[idx.comorbidities_dm] === 'Yes';
    const hasHT = row[idx.comorbidities_ht] === 'Yes';
    const hasIHD = row[idx.comorbidities_ihd] === 'Yes';
    const pregnant = pregIdx >= 0 && row[pregIdx] === 'Yes';
    const hasCKD = ckdIdx >= 0 && row[ckdIdx] === 'Yes';
    const immunosuppressed = immIdx >= 0 && row[immIdx] === 'Yes';
    const onNsaids = row[idx.nsaids] === 'Yes';
    const onSteroids = row[idx.steroids] === 'Yes';
    
    const isHighRisk = hasBleeding || hasComorbidity || lateAdmission || age > 60 || hasDM || hasHT || hasIHD || pregnant || hasCKD || immunosuppressed || onNsaids || onSteroids;
    
    if (isHighRisk) {
      const meds = [];
      if (onNsaids) meds.push('NSAID');
      if (row[idx.antiplatelet] === 'Yes') meds.push('Antiplatelet');
      if (onSteroids) meds.push('Steroid');
      
      const factors = [];
      if (hasDM) factors.push('Diabetes');
      if (hasHT) factors.push('HT');
      if (hasIHD) factors.push('IHD');
      if (pregnant) factors.push('Pregnant');
      if (hasCKD) factors.push('CKD');
      if (immunosuppressed) factors.push('Immuno');
      if (onNsaids) factors.push('NSAID');
      if (onSteroids) factors.push('Steroid');
      if (age > 60) factors.push('Elderly');
      if (hasBleeding) factors.push('Bleeding');
      
      highRisk.push({
        id: row[idx.patientId] || ('P' + (i + 1)),
        patientId: row[idx.patientId] || ('Patient ' + (i + 1)),
        index: i,
        age: age,
        hasBleeding: hasBleeding,
        hasComorbidity: hasComorbidity,
        hasDM: hasDM,
        hasHT: hasHT,
        hasIHD: hasIHD,
        pregnant: pregnant,
        hasCKD: hasCKD,
        immunosuppressed: immunosuppressed,
        onNsaids: onNsaids,
        onSteroids: onSteroids,
        medications: meds.length > 0 ? meds.join(', ') : 'None',
        riskFactors: factors.join(', '),
        reason: factors.slice(0, 2).join(', ')
      });
    }
  });
  
  return highRisk;
}

function getRiskReason(row, idx) {
  const reasons = [];
  if (row[idx.bleeding] === 'Yes') reasons.push('Bleeding');
  if (row[idx.comorbidities_dm] === 'Yes') reasons.push('Diabetes');
  if (row[idx.comorbidities_ht] === 'Yes') reasons.push('Hypertension');
  const age = parseFloat(row[idx.age]) || 0;
  if (age > 65) reasons.push('Elderly');
  if (parseFloat(row[idx.dayOfIllness]) > 5) reasons.push('Late admission');
  return reasons.join(', ') || 'Multiple risk factors';
}

function generateClinicalInsights(data, idx) {
  const insights = [];
  
  const avgAge = data.reduce(function(sum, r) { return sum + (parseFloat(r[idx.age]) || 0); }, 0) / (data.length || 1);
  const bleedingRate = (data.filter(function(r) { return r[idx.bleeding] === 'Yes'; }).length / (data.length || 1)) * 100;
  const comorbidityRate = (data.filter(function(r) { 
    return r[idx.comorbidities_ht] === 'Yes' || r[idx.comorbidities_dm] === 'Yes';
  }).length / (data.length || 1)) * 100;
  const avgDayAdmitted = data.reduce(function(sum, r) { return sum + (parseFloat(r[idx.dayOfIllness]) || 0); }, 0) / (data.length || 1);
  const femaleCount = data.filter(function(r) { return (r[idx.gender] || '').toLowerCase() === 'female'; }).length;
  
  insights.push({
    type: 'demographic',
    icon: '👥',
    title: 'Patient Demographics',
    text: 'Average patient age is ' + avgAge.toFixed(1) + ' years with ' + (femaleCount > data.length / 2 ? 'female' : 'male') + ' predominance.'
  });
  
  if (bleedingRate > 15) {
    insights.push({
      type: 'warning',
      icon: '⚠️',
      title: 'Elevated Bleeding Rate',
      text: bleedingRate.toFixed(1) + '% of patients show bleeding manifestations - monitor closely.'
    });
  }
  
  if (comorbidityRate > 30) {
    insights.push({
      type: 'alert',
      icon: '🔴',
      title: 'High Comorbidity Burden',
      text: comorbidityRate.toFixed(1) + '% have comorbidities - higher risk of complications.'
    });
  }
  
  if (avgDayAdmitted > 4) {
    insights.push({
      type: 'timing',
      icon: '⏰',
      title: 'Late Presentation',
      text: 'Average admission on day ' + avgDayAdmitted.toFixed(1) + ' - early presentation improves outcomes.'
    });
  }
  
  const ageGroups = { young: 0, adult: 0, elderly: 0 };
  data.forEach(function(r) {
    const age = parseFloat(r[idx.age]) || 0;
    if (age < 20) ageGroups.young++;
    else if (age < 60) ageGroups.adult++;
    else ageGroups.elderly++;
  });
  
  if (data.length > 0) {
    insights.push({
      type: 'trend',
      icon: '📊',
      title: 'Age Distribution',
      text: ((ageGroups.adult / data.length) * 100).toFixed(0) + '% of cases in productive adults (20-59 years).'
    });
  }
  
  return insights;
}

function analyzeComorbidities(data, idx) {
  return {
    hypertension: data.filter(function(r) { return r[idx.comorbidities_ht] === 'Yes'; }).length,
    diabetes: data.filter(function(r) { return r[idx.comorbidities_dm] === 'Yes'; }).length,
    ischaemicHeart: data.filter(function(r) { return r[idx.comorbidities_ihd] === 'Yes'; }).length,
    asthma: data.filter(function(r) { return r[idx.comorbidities_asthma] === 'Yes'; }).length,
    chronicKidney: data.filter(function(r) { return r[idx.comorbidities_ckd] === 'Yes'; }).length,
    chronicLiver: data.filter(function(r) { return r[idx.comorbidities_cld] === 'Yes'; }).length
  };
}

function analyzeMedications(data, idx) {
  return {
    nsaids: data.filter(function(r) { return r[idx.nsaids] === 'Yes'; }).length,
    antiplatelet: data.filter(function(r) { return r[idx.antiplatelet] === 'Yes'; }).length,
    steroids: data.filter(function(r) { return r[idx.steroids] === 'Yes'; }).length
  };
}

function calculateCohortComparison(data, idx) {
  const cohorts = {
    mild: { count: 0, avgAge: 0, avgBmi: 0, comorbidities: 0, bleeding: 0 },
    moderate: { count: 0, avgAge: 0, avgBmi: 0, comorbidities: 0, bleeding: 0 },
    severe: { count: 0, avgAge: 0, avgBmi: 0, comorbidities: 0, bleeding: 0 }
  };
  
  data.forEach(function(row) {
    const severity = assessSeverity(row, idx);
    const cohort = cohorts[severity.toLowerCase()];
    
    cohort.count++;
    cohort.avgAge += parseFloat(row[idx.age]) || 0;
    cohort.avgBmi += parseFloat(row[idx.bmi]) || 0;
    if (row[idx.comorbidities_ht] === 'Yes' || row[idx.comorbidities_dm] === 'Yes') cohort.comorbidities++;
    if (row[idx.bleeding] === 'Yes') cohort.bleeding++;
  });
  
  ['mild', 'moderate', 'severe'].forEach(function(severity) {
    const c = cohorts[severity];
    const n = c.count || 1;
    c.avgAge = (c.avgAge / n).toFixed(1);
    c.avgBmi = (c.avgBmi / n).toFixed(1);
    c.comorbidityRate = ((c.comorbidities / n) * 100).toFixed(1);
    c.bleedingRate = ((c.bleeding / n) * 100).toFixed(1);
  });
  
  return cohorts;
}

function refreshData() {
  cache.remove('dashboardData');
  cache.remove('kpisOnly');
  cache.remove('collMonitorFresh');
  return { status: 'success', message: 'Cache cleared. Data will refresh on next load.' };
}

function calculateCohortAnalysis(data, headers, idx) {
  const cohorts = {
    mild: { count: 0, totalAge: 0, totalBmi: 0, comorbidities: 0, bleeding: 0, hypertension: 0, diabetes: 0, ihd: 0, ckd: 0, immuno: 0, nsaids: 0, antiplatelet: 0, steroids: 0, elderly: 0, lateAdmission: 0, male: 0, female: 0, totalDay: 0, pregnant: 0 },
    moderate: { count: 0, totalAge: 0, totalBmi: 0, comorbidities: 0, bleeding: 0, hypertension: 0, diabetes: 0, ihd: 0, ckd: 0, immuno: 0, nsaids: 0, antiplatelet: 0, steroids: 0, elderly: 0, lateAdmission: 0, male: 0, female: 0, totalDay: 0, pregnant: 0 },
    severe: { count: 0, totalAge: 0, totalBmi: 0, comorbidities: 0, bleeding: 0, hypertension: 0, diabetes: 0, ihd: 0, ckd: 0, immuno: 0, nsaids: 0, antiplatelet: 0, steroids: 0, elderly: 0, lateAdmission: 0, male: 0, female: 0, totalDay: 0, pregnant: 0 }
  };
  
  const pregIdx = headers.findIndex(function(h) { return h.toLowerCase().includes('pregnancy'); });
  const ckdIdx = headers.findIndex(function(h) { return h.toLowerCase().includes('chronic kidney'); });
  const immIdx = headers.findIndex(function(h) { return h.toLowerCase().includes('immunosuppressive'); });
  
  data.forEach(function(row) {
    const severity = assessSeverity(row, idx);
    const c = cohorts[severity.toLowerCase()];
    
    const age = parseFloat(row[idx.age]) || 0;
    const bmi = parseFloat(row[idx.bmi]) || 0;
    const day = parseFloat(row[idx.dayOfIllness]) || 0;
    const gender = (row[idx.gender] || '').toLowerCase();
    
    c.count++;
    c.totalAge += age;
    c.totalBmi += bmi;
    c.totalDay += day;
    
    if (gender === 'male') c.male++;
    else if (gender === 'female') c.female++;
    
    if (row[idx.comorbidities_ht] === 'Yes') { c.comorbidities++; c.hypertension++; }
    if (row[idx.comorbidities_dm] === 'Yes') { c.comorbidities++; c.diabetes++; }
    if (row[idx.comorbidities_ihd] === 'Yes') { c.comorbidities++; c.ihd++; }
    if (ckdIdx >= 0 && row[ckdIdx] === 'Yes') { c.comorbidities++; c.ckd++; }
    if (immIdx >= 0 && row[immIdx] === 'Yes') { c.comorbidities++; c.immuno++; }
    if (pregIdx >= 0 && row[pregIdx] === 'Yes') { c.pregnant++; }
    
    if (row[idx.bleeding] === 'Yes') c.bleeding++;
    if (row[idx.nsaids] === 'Yes') c.nsaids++;
    if (row[idx.antiplatelet] === 'Yes') c.antiplatelet++;
    if (row[idx.steroids] === 'Yes') c.steroids++;
    if (age > 60) c.elderly++;
    if (day > 5) c.lateAdmission++;
  });
  
  var n = data.length || 1;
  var keyFindings = [];
  
  ['mild', 'moderate', 'severe'].forEach(function(sev) {
    var c = cohorts[sev];
    var cnt = c.count || 1;
    c.avgAge = (c.totalAge / cnt).toFixed(1);
    c.avgBmi = (c.totalBmi / cnt).toFixed(1);
    c.avgDay = (c.totalDay / cnt).toFixed(1);
    c.comorbidityRate = ((c.comorbidities / cnt) * 100).toFixed(1);
    c.bleedingRate = ((c.bleeding / cnt) * 100).toFixed(1);
    c.malePct = ((c.male / cnt) * 100).toFixed(1);
    c.femalePct = ((c.female / cnt) * 100).toFixed(1);
    c.nsaidsPct = ((c.nsaids / cnt) * 100).toFixed(1);
    c.antiplateletPct = ((c.antiplatelet / cnt) * 100).toFixed(1);
    c.steroidsPct = ((c.steroids / cnt) * 100).toFixed(1);
    
    delete c.totalAge;
    delete c.totalBmi;
    delete c.totalDay;
  });
  
  var severeCohort = cohorts.severe;
  var moderateCohort = cohorts.moderate;
  var mildCohort = cohorts.mild;
  
  var severeComorbidityRate = parseFloat(severeCohort.comorbidityRate) || 0;
  var moderateComorbidityRate = parseFloat(moderateCohort.comorbidityRate) || 0;
  var mildComorbidityRate = parseFloat(mildCohort.comorbidityRate) || 0;
  
  if (severeComorbidityRate > mildComorbidityRate * 1.5) {
    keyFindings.push({
      type: 'alert',
      icon: '🔴',
      title: 'Severity-Comorbidity Link',
      text: 'Severe cases show ' + severeComorbidityRate + '% comorbidity rate vs ' + mildComorbidityRate + '% in mild cases - strong correlation'
    });
  }
  
  if (parseFloat(severeCohort.bleedingRate) > 30) {
    keyFindings.push({
      type: 'alert',
      icon: '🩸',
      title: 'Bleeding in Severe Cases',
      text: severeCohort.bleedingRate + '% of severe cases have bleeding manifestations'
    });
  }
  
  if (parseFloat(severeCohort.avgAge) > parseFloat(mildCohort.avgAge) + 10) {
    keyFindings.push({
      type: 'warning',
      icon: '👴',
      title: 'Age-Severity Association',
      text: 'Severe cases avg ' + severeCohort.avgAge + ' yrs vs ' + mildCohort.avgAge + ' yrs in mild - older patients at higher risk'
    });
  }
  
  if (parseFloat(severeCohort.nsaidsPct) > parseFloat(mildCohort.nsaidsPct)) {
    keyFindings.push({
      type: 'warning',
      icon: '💊',
      title: 'NSAID Pattern',
      text: 'NSAID use: ' + severeCohort.nsaidsPct + '% severe vs ' + mildCohort.nsaidsPct + '% mild - possible association with severity'
    });
  }
  
  if (parseFloat(severeCohort.lateAdmission) / (severeCohort.count || 1) > 0.3) {
    keyFindings.push({
      type: 'warning',
      icon: '⏰',
      title: 'Late Presentation',
      text: Math.round(parseFloat(severeCohort.lateAdmission)) + ' severe cases admitted after day 5 - delayed care increases risk'
    });
  }
  
  var highestComorbidity = 'Mild';
  var highestComorbidityRate = mildComorbidityRate;
  if (moderateComorbidityRate > highestComorbidityRate) { highestComorbidity = 'Moderate'; highestComorbidityRate = moderateComorbidityRate; }
  if (severeComorbidityRate > highestComorbidityRate) { highestComorbidity = 'Severe'; highestComorbidityRate = severeComorbidityRate; }
  
  var highestBleeding = 'Mild';
  var highestBleedingRate = parseFloat(mildCohort.bleedingRate) || 0;
  if (parseFloat(moderateCohort.bleedingRate) > highestBleedingRate) { highestBleeding = 'Moderate'; highestBleedingRate = parseFloat(moderateCohort.bleedingRate); }
  if (parseFloat(severeCohort.bleedingRate) > highestBleedingRate) { highestBleeding = 'Severe'; highestBleedingRate = parseFloat(severeCohort.bleedingRate); }
  
  var outcomeIndicators = [];
  
  if (severeCohort.count > 0) {
    outcomeIndicators.push({
      icon: '📊',
      title: 'Severity Distribution',
      text: 'Mild: ' + mildCohort.count + ' (' + ((mildCohort.count / n) * 100).toFixed(1) + '%), Moderate: ' + moderateCohort.count + ' (' + ((moderateCohort.count / n) * 100).toFixed(1) + '%), Severe: ' + severeCohort.count + ' (' + ((severeCohort.count / n) * 100).toFixed(1) + '%)'
    });
  }
  
  if (parseFloat(severeCohort.avgDay) > 4) {
    outcomeIndicators.push({
      type: 'warning',
      icon: '⏰',
      title: 'Admission Timing',
      text: 'Severe cases average admission on day ' + severeCohort.avgDay + ' - earlier intervention could improve outcomes'
    });
  }
  
  if (parseFloat(severeCohort.comorbidityRate) > 50) {
    outcomeIndicators.push({
      type: 'alert',
      icon: '🩺',
      title: 'High-Risk Population',
      text: severeCohort.comorbidityRate + '% of severe cases have comorbidities - proactive monitoring recommended'
    });
  }
  
  return {
    cohorts: cohorts,
    keyFindings: keyFindings,
    highestComorbidity: highestComorbidity,
    highestComorbidityRate: highestComorbidityRate,
    highestBleeding: highestBleeding,
    highestBleedingRate: highestBleedingRate,
    outcomeIndicators: outcomeIndicators
  };
}

function exportReport() {
  const data = getDashboardData();
  let report = 'CSTH Dengue Research Report\n';
  report += 'Generated: ' + new Date().toLocaleString() + '\n';
  report += '=====================================\n\n';
  report += 'SUMMARY\n';
  report += '-------\n';
  report += 'Total Patients: ' + data.kpis.totalPatients + '\n';
  report += 'Severity Distribution: ' + data.severityDistribution.counts.join(' / ') + ' (Mild / Moderate / Severe)\n';
  report += 'Average Age: ' + data.kpis.averageAge + ' years\n';
  report += 'Average Day of Admission: Day ' + data.kpis.averageDayAdmitted + '\n\n';
  report += 'CLINICAL INSIGHTS\n';
  report += '-----------------\n';
  data.clinicalInsights.forEach(function(i) {
    report += '* ' + i.icon + ' ' + i.title + ': ' + i.text + '\n';
  });
  report += '\nHIGH RISK PATIENTS\n';
  report += '------------------\n';
  report += data.highRiskPatients.length + ' patients identified\n';
  data.highRiskPatients.slice(0, 5).forEach(function(p) {
    report += '* ' + p.patientId + ' (Age: ' + p.age + ') - ' + p.reason + '\n';
  });
  return report;
}

function calculatePatientData(data, headers, idx) {
  const n = data.length || 1;
  let totalAge = 0, comorbidity = 0, medication = 0, pregnant = 0;
  let hypertension = 0, diabetes = 0, ihd = 0, asthma = 0;
  let ckd = 0, cld = 0, immunosuppressed = 0, fattyLiver = 0;
  let priorDengue = 0, onNsaids = 0, onSteroids = 0, onAntiplatelet = 0;
  
  data.forEach(function(row) {
    totalAge += parseFloat(row[idx.age]) || 0;
    
    if (row[idx.comorbidities_ht] === 'Yes') { comorbidity++; hypertension++; }
    if (row[idx.comorbidities_dm] === 'Yes') { comorbidity++; diabetes++; }
    if (row[idx.comorbidities_ihd] === 'Yes') { comorbidity++; ihd++; }
    
    const astIdx = headers.findIndex(function(h) { return h.toLowerCase().includes('asthma'); });
    const ckdIdx = headers.findIndex(function(h) { return h.toLowerCase().includes('chronic kidney'); });
    const cldIdx = headers.findIndex(function(h) { return h.toLowerCase().includes('chronic liver'); });
    const immIdx = headers.findIndex(function(h) { return h.toLowerCase().includes('immunosuppressive'); });
    const flIdx = headers.findIndex(function(h) { return h.toLowerCase().includes('fatty liver'); });
    const pregIdx = headers.findIndex(function(h) { return h.toLowerCase().includes('pregnancy'); });
    
    if (astIdx >= 0 && row[astIdx] === 'Yes') { comorbidity++; asthma++; }
    if (ckdIdx >= 0 && row[ckdIdx] === 'Yes') { comorbidity++; ckd++; }
    if (cldIdx >= 0 && row[cldIdx] === 'Yes') { comorbidity++; cld++; }
    if (immIdx >= 0 && row[immIdx] === 'Yes') { comorbidity++; immunosuppressed++; }
    if (flIdx >= 0 && row[flIdx] === 'Yes') { comorbidity++; fattyLiver++; }
    if (pregIdx >= 0 && row[pregIdx] === 'Yes') { comorbidity++; pregnant++; }
    
    if (row[idx.nsaids] === 'Yes') { medication++; onNsaids++; }
    if (row[idx.steroids] === 'Yes') { medication++; onSteroids++; }
    if (row[idx.antiplatelet] === 'Yes') { medication++; onAntiplatelet++; }
  });
  
  let filledFields = 0;
  data.forEach(function(row) {
    row.forEach(function(cell) {
      if (cell !== '' && cell !== null) filledFields++;
    });
  });
  const avgCompleteness = Math.round((filledFields / (data.length * headers.length)) * 100);
  
  return {
    total: n,
    avgAge: (totalAge / n).toFixed(1),
    comorbidity: comorbidity,
    comorbidityPct: ((comorbidity / n) * 100).toFixed(0),
    medication: medication,
    medicationPct: ((medication / n) * 100).toFixed(0),
    pregnant: pregnant,
    hypertension: hypertension,
    diabetes: diabetes,
    ihd: ihd,
    asthma: asthma,
    ckd: ckd,
    cld: cld,
    immunosuppressed: immunosuppressed,
    fattyLiver: fattyLiver,
    onNsaids: onNsaids,
    onSteroids: onSteroids,
    onAntiplatelet: onAntiplatelet,
    avgCompleteness: avgCompleteness
  };
}

function calculateCollectionProgress(data, headers) {
  if (data.length === 0) return [];
  
  const criticalFields = ['Patient ID', 'Age', 'Gender', 'BMI', 'Date', 'Hypertension', 'Diabetes', 'Smoking', 'Bleeding', 'NSAIDs'];
  const progress = [];
  
  headers.forEach(function(h, i) {
    if (criticalFields.some(function(cf) { return h.toLowerCase().includes(cf.toLowerCase()); })) {
      const filled = data.filter(function(r) { return r[i] !== '' && r[i] !== null; }).length;
      const pct = Math.round((filled / data.length) * 100);
      let level = 'excellent';
      if (pct < 70) level = 'poor';
      else if (pct < 85) level = 'moderate';
      else if (pct < 95) level = 'good';
      
      progress.push({ field: h, pct: pct, level: level });
    }
  });
  
  return progress.sort(function(a, b) { return a.pct - b.pct; });
}

function calculateRiskAnalysis(data, headers, idx) {
  const n = data.length || 1;
  const pregIdx = headers.findIndex(function(h) { return h.toLowerCase().includes('pregnancy'); });
  const ckdIdx = headers.findIndex(function(h) { return h.toLowerCase().includes('chronic kidney'); });
  const immIdx = headers.findIndex(function(h) { return h.toLowerCase().includes('immunosuppressive'); });
  const flIdx = headers.findIndex(function(h) { return h.toLowerCase().includes('fatty liver'); });
  const smokingIdx = headers.findIndex(function(h) { return h.toLowerCase().includes('smoking'); });
  const alcoholIdx = headers.findIndex(function(h) { return h.toLowerCase().includes('alcohol'); });
  
  let nsaidsCount = 0, bleedingCount = 0, elderlyCount = 0, diabetesCount = 0;
  let hypertensionCount = 0, ihdCount = 0, pregnantCount = 0, ckdCount = 0;
  let immunoCount = 0, lateAdmissionCount = 0, highRiskCount = 0;
  let steroidCount = 0, antiplateletCount = 0, fattyLiverCount = 0;
  let smokingCount = 0, alcoholCount = 0;
  let totalCollectionScore = 0;
  const collectionFields = ['Patient ID', 'Age', 'Gender', 'BMI', 'Hypertension', 'Diabetes', 'Bleeding', 'NSAIDs', 'Smoking', 'Day of illness'];
  
  data.forEach(function(row, i) {
    let rowScore = 0;
    
    if (row[idx.nsaids] === 'Yes') { nsaidsCount++; rowScore++; }
    if (row[idx.antiplatelet] === 'Yes') { antiplateletCount++; rowScore++; }
    if (row[idx.steroids] === 'Yes') { steroidCount++; rowScore++; }
    if (row[idx.bleeding] === 'Yes') { bleedingCount++; rowScore += 2; }
    
    const age = parseFloat(row[idx.age]) || 0;
    if (age > 60) { elderlyCount++; rowScore += 2; }
    
    if (row[idx.comorbidities_dm] === 'Yes') { diabetesCount++; rowScore++; }
    if (row[idx.comorbidities_ht] === 'Yes') { hypertensionCount++; rowScore++; }
    if (row[idx.comorbidities_ihd] === 'Yes') { ihdCount++; rowScore += 2; }
    if (pregIdx >= 0 && row[pregIdx] === 'Yes') { pregnantCount++; rowScore += 2; }
    if (ckdIdx >= 0 && row[ckdIdx] === 'Yes') { ckdCount++; rowScore += 2; }
    if (immIdx >= 0 && row[immIdx] === 'Yes') { immunoCount++; rowScore += 2; }
    if (flIdx >= 0 && row[flIdx] === 'Yes') { fattyLiverCount++; }
    
    const day = parseFloat(row[idx.dayOfIllness]) || 0;
    if (day > 5) { lateAdmissionCount++; rowScore++; }
    
    if (smokingIdx >= 0 && (row[smokingIdx] || '').toString().toLowerCase().includes('current')) { smokingCount++; }
    if (alcoholIdx >= 0 && row[alcoholIdx] === 'Yes') { alcoholCount++; }
    
    if (rowScore >= 3) highRiskCount++;
    totalCollectionScore += rowScore;
  });
  
  const avgCollectionRate = Math.round((totalCollectionScore / (n * 10)) * 100);
  
  const collectionProgressByRisk = [];
  const riskFields = ['NSAIDs', 'Age', 'Bleeding', 'Hypertension', 'Diabetes', 'Day of illness'];
  headers.forEach(function(h, i) {
    if (riskFields.some(function(rf) { return h.toLowerCase().includes(rf.toLowerCase()); })) {
      const filled = data.filter(function(r) { return r[i] !== '' && r[i] !== null; }).length;
      const pct = Math.round((filled / n) * 100);
      let level = 'excellent';
      if (pct < 70) level = 'poor';
      else if (pct < 85) level = 'moderate';
      else if (pct < 95) level = 'good';
      collectionProgressByRisk.push({ field: h, pct: pct, level: level });
    }
  });
  
  const clinicalInsights = [];
  const insightsArr = generateClinicalInsights(data, idx);
  clinicalInsights.push.apply(clinicalInsights, insightsArr);
  
  if (nsaidsCount > 0) {
    clinicalInsights.push({
      type: 'alert',
      icon: '💊',
      title: 'NSAID Risk Alert',
      text: nsaidsCount + ' patients (' + ((nsaidsCount/n)*100).toFixed(1) + '%) on NSAIDs - associated with increased bleeding risk'
    });
  }
  
  if (elderlyCount > 0) {
    clinicalInsights.push({
      type: 'warning',
      icon: '👴',
      title: 'Elderly Patients',
      text: elderlyCount + ' patients over 60 years - higher risk of complications'
    });
  }
  
  if (pregnantCount > 0) {
    clinicalInsights.push({
      type: 'alert',
      icon: '🤰',
      title: 'Pregnancy Risk',
      text: pregnantCount + ' pregnant patients - requires special monitoring and care'
    });
  }
  
  if (ckdCount > 0 || immunoCount > 0) {
    clinicalInsights.push({
      type: 'alert',
      icon: '🛡️',
      title: 'High-Risk Conditions',
      text: ckdCount + ' CKD + ' + immunoCount + ' immunosuppressed patients - fluid management critical'
    });
  }
  
  if (lateAdmissionCount > 0) {
    clinicalInsights.push({
      type: 'warning',
      icon: '⏰',
      title: 'Late Admission',
      text: lateAdmissionCount + ' patients admitted after day 5 - outcomes may be affected'
    });
  }
  
  if (bleedingCount > 0) {
    clinicalInsights.push({
      type: 'alert',
      icon: '🩸',
      title: 'Bleeding Manifestations',
      text: bleedingCount + ' patients (' + ((bleedingCount/n)*100).toFixed(1) + '%) with bleeding - monitor closely'
    });
  }
  
  if (highRiskCount > 0) {
    clinicalInsights.push({
      type: 'alert',
      icon: '🚨',
      title: 'High-Risk Count',
      text: highRiskCount + ' patients with multiple risk factors - priority monitoring needed'
    });
  }
  
  if (avgCollectionRate < 80) {
    clinicalInsights.push({
      type: 'warning',
      icon: '📝',
      title: 'Data Collection Gap',
      text: avgCollectionRate + '% risk factor documentation - improve data collection'
    });
  }
  
  return {
    nsaidsCount: nsaidsCount,
    bleedingCount: bleedingCount,
    elderlyCount: elderlyCount,
    diabetesCount: diabetesCount,
    hypertensionCount: hypertensionCount,
    ihdCount: ihdCount,
    pregnantCount: pregnantCount,
    ckdCount: ckdCount,
    immunoCount: immunoCount,
    lateAdmissionCount: lateAdmissionCount,
    highRiskCount: highRiskCount,
    steroidCount: steroidCount,
    antiplateletCount: antiplateletCount,
    fattyLiverCount: fattyLiverCount,
    smokingCount: smokingCount,
    alcoholCount: alcoholCount,
    avgCollectionRate: avgCollectionRate,
    collectionProgressByRisk: collectionProgressByRisk,
    clinicalInsights: clinicalInsights
  };
}

function calculateWarningSigns(data, headers, idx) {
  var n = data.length || 1;
  var abdominalPain = 0, vomiting = 0, bleeding = 0, lethargy = 0, hctRise = 0, fluidAccum = 0;

  // Section 15 warning sign columns (post-fever deterioration)
  var ws15_5Idx = headers.findIndex(function(h) { return h.includes('15.5'); }); // Abdominal pain/tenderness

  data.forEach(function(row) {
    // Abdominal pain: prefer section 15.5 warning sign, fallback to section 14.9
    var hasAbdPain = (ws15_5Idx >= 0 && row[ws15_5Idx] === 'Yes') ||
                     (idx.abdominalPain >= 0 && row[idx.abdominalPain] === 'Yes');
    if (hasAbdPain) abdominalPain++;

    // Persistent vomiting: prefer section 15.3, fallback to section 14.8
    var hasPersistVomit = (idx.persistentVomiting >= 0 && row[idx.persistentVomiting] === 'Yes') ||
                          (idx.vomiting >= 0 && row[idx.vomiting] === 'Yes');
    if (hasPersistVomit) vomiting++;

    // Mucosal bleeding: section 17.2, fallback to bleeding manifestations
    var hasBleeding = (idx.mucosalBleeding >= 0 && row[idx.mucosalBleeding] === 'Yes') ||
                      (idx.bleeding >= 0 && row[idx.bleeding] === 'Yes');
    if (hasBleeding) bleeding++;

    // Lethargy/restlessness: section 15.4 (was incorrectly using nausea)
    if (idx.lethargy >= 0 && row[idx.lethargy] === 'Yes') lethargy++;

    // Rapid HCT rise: section 17.4
    if (idx.hctRiseFlag >= 0 && row[idx.hctRiseFlag] === 'Yes') hctRise++;

    // Clinically detected fluid accumulation: section 17.1
    if (idx.fluidAccumulation >= 0 && row[idx.fluidAccumulation] === 'Yes') fluidAccum++;
  });

  return {
    abdominalPain:    abdominalPain,
    abdominalPainPct: ((abdominalPain / n) * 100).toFixed(1),
    vomiting:         vomiting,
    vomitingPct:      ((vomiting / n) * 100).toFixed(1),
    bleeding:         bleeding,
    bleedingPct:      ((bleeding / n) * 100).toFixed(1),
    lethargy:         lethargy,
    lethargyPct:      ((lethargy / n) * 100).toFixed(1),
    hctRise:          hctRise,
    hctRisePct:       ((hctRise / n) * 100).toFixed(1),
    fluidAccumulation:    fluidAccum,
    fluidAccumulationPct: ((fluidAccum / n) * 100).toFixed(1)
  };
}

function calculateCriticalPhase(data, idx) {
  var inCriticalPhase = 0, earlyPhase = 0, recoveryPhase = 0;
  var severeInCritical = 0;
  
  data.forEach(function(row) {
    var day = parseFloat(row[idx.dayOfIllness]) || 0;
    var severity = assessSeverity(row, idx);
    
    if (day >= 4 && day <= 6) {
      inCriticalPhase++;
      if (severity === 'Severe') severeInCritical++;
    } else if (day < 4) {
      earlyPhase++;
    } else {
      recoveryPhase++;
    }
  });
  
  return {
    inCriticalPhase: inCriticalPhase,
    earlyPhase: earlyPhase,
    recoveryPhase: recoveryPhase,
    severeInCritical: severeInCritical
  };
}

function calculateSymptomPatterns(data, headers, idx) {
  var n = data.length || 1;
  var fever = 0, headache = 0, retroOrbital = 0, myalgia = 0;
  var arthralgia = 0, rash = 0, nausea = 0, vomiting = 0, abdominalPain = 0;
  var classicTriad = 0;
  
  var abdominalIdx = headers.findIndex(function(h) { return h.toLowerCase().includes('abdominal pain'); });
  
  data.forEach(function(row) {
    if (row[idx.fever] === 'Yes') fever++;
    if (row[idx.headache] === 'Yes') headache++;
    if (row[idx.retroOrbital] === 'Yes') retroOrbital++;
    if (row[idx.myalgia] === 'Yes') myalgia++;
    if (row[idx.arthralgia] === 'Yes') arthralgia++;
    if (row[idx.rash] === 'Yes') rash++;
    if (row[idx.nausea] === 'Yes') nausea++;
    if (row[idx.vomiting] === 'Yes') vomiting++;
    if (row[idx.abdominalPain] === 'Yes' || (abdominalIdx >= 0 && row[abdominalIdx] === 'Yes')) abdominalPain++;
    
    if (row[idx.fever] === 'Yes' && row[idx.headache] === 'Yes' && row[idx.rash] === 'Yes') {
      classicTriad++;
    }
  });
  
  return {
    total: n,
    fever: fever,
    feverPct: ((fever / n) * 100).toFixed(1),
    headache: headache,
    headachePct: ((headache / n) * 100).toFixed(1),
    retroOrbital: retroOrbital,
    retroOrbitalPct: ((retroOrbital / n) * 100).toFixed(1),
    myalgia: myalgia,
    myalgiaPct: ((myalgia / n) * 100).toFixed(1),
    arthralgia: arthralgia,
    arthralgiaPct: ((arthralgia / n) * 100).toFixed(1),
    rash: rash,
    rashPct: ((rash / n) * 100).toFixed(1),
    nausea: nausea,
    nauseaPct: ((nausea / n) * 100).toFixed(1),
    vomiting: vomiting,
    vomitingPct: ((vomiting / n) * 100).toFixed(1),
    abdominalPain: abdominalPain,
    abdominalPainPct: ((abdominalPain / n) * 100).toFixed(1),
    classicTriad: classicTriad,
    classicTriadPct: ((classicTriad / n) * 100).toFixed(1)
  };
}

// ══════════════════════════════════════════════════════════════════
//  LAB VALUES
// ══════════════════════════════════════════════════════════════════
function calculateLabValues(data, idx) {
  var n = data.length || 1;
  var plateletAdmValues = [], plateletLowValues = [];
  var wbcAdmValues = [];
  var altHighValues = [], astHighValues = [];
  var creatinineValues = [], albuminValues = [], inrValues = [];

  var plateletCritical = 0, plateletLow20 = 0, plateletLow50 = 0, plateletLow100 = 0;
  var altElevated = 0, astElevated = 0;
  var creatinineElevated = 0, albuminLow = 0, inrElevated = 0;

  data.forEach(function(row) {
    var platAdm = parseFloat(row[idx.plateletAdm]) || 0;
    var platLow = parseFloat(row[idx.plateletLow]) || 0;
    var wbc     = parseFloat(row[idx.wbcAdm])     || 0;
    var altH    = parseFloat(row[idx.altHigh])    || 0;
    var astH    = parseFloat(row[idx.astHigh])    || 0;
    var creat   = parseFloat(row[idx.creatinine]) || 0;
    var alb     = parseFloat(row[idx.albumin])    || 0;
    var inr     = parseFloat(row[idx.inr])        || 0;

    if (platAdm > 0) plateletAdmValues.push(platAdm);
    if (platLow > 0) {
      plateletLowValues.push(platLow);
      if (platLow < 20)       { plateletCritical++; plateletLow20++; }
      else if (platLow < 50)  plateletLow50++;
      else if (platLow < 100) plateletLow100++;
    }
    if (wbc   > 0) wbcAdmValues.push(wbc);
    if (altH  > 0) { altHighValues.push(altH);  if (altH  > 40)  altElevated++; }
    if (astH  > 0) { astHighValues.push(astH);  if (astH  > 40)  astElevated++; }
    if (creat > 0) { creatinineValues.push(creat); if (creat > 1.2) creatinineElevated++; }
    if (alb   > 0) { albuminValues.push(alb);   if (alb   < 35)  albuminLow++; }
    if (inr   > 0) { inrValues.push(inr);       if (inr   > 1.5) inrElevated++; }
  });

  function avg(arr)    { return arr.length ? (arr.reduce(function(a,b){return a+b;},0)/arr.length).toFixed(1) : 'N/A'; }
  function minVal(arr) { return arr.length ? Math.min.apply(null,arr).toFixed(1) : 'N/A'; }
  function maxVal(arr) { return arr.length ? Math.max.apply(null,arr).toFixed(1) : 'N/A'; }

  return {
    plateletAdmAvg:  avg(plateletAdmValues),
    plateletLowAvg:  avg(plateletLowValues),
    plateletLowMin:  minVal(plateletLowValues),
    plateletCritical: plateletCritical,
    plateletLow20:   plateletLow20,
    plateletLow50:   plateletLow50,
    plateletLow100:  plateletLow100,
    plateletCount:   plateletLowValues.length,
    wbcAdmAvg:       avg(wbcAdmValues),
    wbcCount:        wbcAdmValues.length,
    altHighAvg:      avg(altHighValues),
    altMax:          maxVal(altHighValues),
    altElevated:     altElevated,
    astHighAvg:      avg(astHighValues),
    astMax:          maxVal(astHighValues),
    astElevated:     astElevated,
    creatinineAvg:   avg(creatinineValues),
    creatinineMax:   maxVal(creatinineValues),
    creatinineElevated: creatinineElevated,
    albuminAvg:      avg(albuminValues),
    albuminMin:      minVal(albuminValues),
    albuminLow:      albuminLow,
    inrAvg:          avg(inrValues),
    inrMax:          maxVal(inrValues),
    inrElevated:     inrElevated
  };
}

// ══════════════════════════════════════════════════════════════════
//  SEROLOGY
// ══════════════════════════════════════════════════════════════════
function calculateSerology(data, idx) {
  var n = data.length || 1;
  var ns1Pos = 0, ns1Neg = 0;
  var iggPos = 0, iggNeg = 0;
  var igmPos = 0, igmNeg = 0;

  data.forEach(function(row) {
    var ns1 = (idx.ns1 >= 0 ? row[idx.ns1] : '').toString().toLowerCase();
    var igg = (idx.igg >= 0 ? row[idx.igg] : '').toString().toLowerCase();
    var igm = (idx.igm >= 0 ? row[idx.igm] : '').toString().toLowerCase();

    if (ns1 === 'positive' || ns1.includes('positive')) ns1Pos++;
    else if (ns1 === 'negative' || ns1.includes('negative')) ns1Neg++;

    if (igg === 'positive' || igg.includes('positive')) iggPos++;
    else if (igg === 'negative' || igg.includes('negative')) iggNeg++;

    if (igm === 'positive' || igm.includes('positive')) igmPos++;
    else if (igm === 'negative' || igm.includes('negative')) igmNeg++;
  });

  // Primary: IgM+ and IgG-  |  Secondary: IgG+
  var secondaryInfection = iggPos;
  var primaryInfection   = Math.max(0, igmPos - iggPos);
  var ns1Total           = ns1Pos + ns1Neg || 1;

  return {
    ns1Positive:    ns1Pos,
    ns1Negative:    ns1Neg,
    ns1PositivePct: ((ns1Pos / ns1Total) * 100).toFixed(1),
    iggPositive:    iggPos,
    iggPositivePct: ((iggPos / n) * 100).toFixed(1),
    igmPositive:    igmPos,
    igmPositivePct: ((igmPos / n) * 100).toFixed(1),
    primaryInfection:   primaryInfection,
    secondaryInfection: secondaryInfection
  };
}

// ══════════════════════════════════════════════════════════════════
//  OUTCOMES
// ══════════════════════════════════════════════════════════════════
function calculateOutcomes(data, idx) {
  var n = data.length || 1;
  var icuCount = 0, hduCount = 0;
  var totalHospDays = 0, hospDayCount = 0;
  var diagDistribution = {}, outcomeDistribution = {};
  var compBleeding = 0, compMOD = 0, compARDS = 0, compALF = 0;
  var compAKI = 0, compMyocarditis = 0, compFluidOverload = 0;

  data.forEach(function(row) {
    // ICU / HDU
    if ((idx.icuAdmit >= 0 ? row[idx.icuAdmit] : '').toString().toLowerCase() === 'yes') icuCount++;
    if ((idx.hduAdmit >= 0 ? row[idx.hduAdmit] : '').toString().toLowerCase() === 'yes') hduCount++;

    // Hospital days
    var days = parseFloat(idx.hospitalDays >= 0 ? row[idx.hospitalDays] : 0) || 0;
    if (days > 0) { totalHospDays += days; hospDayCount++; }

    // Final diagnosis distribution
    var diag = (idx.finalDiagnosis >= 0 ? row[idx.finalDiagnosis] : '').toString().trim();
    if (diag) diagDistribution[diag] = (diagDistribution[diag] || 0) + 1;

    // Outcome distribution
    var out = (idx.outcome >= 0 ? row[idx.outcome] : '').toString().trim();
    if (out) outcomeDistribution[out] = (outcomeDistribution[out] || 0) + 1;

    // Complications — columns contain the option text when checked, or blank
    function hasComp(colIdx) { var v = colIdx >= 0 ? row[colIdx] : ''; return v && v !== ''; }
    if (hasComp(idx.compBleeding))     compBleeding++;
    if (hasComp(idx.compMOD))          compMOD++;
    if (hasComp(idx.compARDS))         compARDS++;
    if (hasComp(idx.compALF))          compALF++;
    if (hasComp(idx.compAKI))          compAKI++;
    if (hasComp(idx.compMyocarditis))  compMyocarditis++;
    if (hasComp(idx.compFluidOverload)) compFluidOverload++;
  });

  // Mortality — count "Died" / "Death" outcomes
  var deathCount = 0;
  Object.keys(outcomeDistribution).forEach(function(k) {
    if (k.toLowerCase().includes('died') || k.toLowerCase().includes('death') || k.toLowerCase().includes('expired')) {
      deathCount += outcomeDistribution[k];
    }
  });

  return {
    icuCount:    icuCount,
    icuPct:      ((icuCount / n) * 100).toFixed(1),
    hduCount:    hduCount,
    hduPct:      ((hduCount / n) * 100).toFixed(1),
    avgHospitalDays: hospDayCount > 0 ? (totalHospDays / hospDayCount).toFixed(1) : 'N/A',
    diagDistribution:    diagDistribution,
    outcomeDistribution: outcomeDistribution,
    deathCount:   deathCount,
    mortalityPct: ((deathCount / n) * 100).toFixed(1),
    complications: {
      severebleeding: compBleeding,
      mod:            compMOD,
      ards:           compARDS,
      alf:            compALF,
      aki:            compAKI,
      myocarditis:    compMyocarditis,
      fluidOverload:  compFluidOverload,
      total: compBleeding + compMOD + compARDS + compALF + compAKI + compMyocarditis + compFluidOverload
    }
  };
}

// ══════════════════════════════════════════════════════════════════
//  COLLECTION MONITOR (enhanced — real-time)
// ══════════════════════════════════════════════════════════════════
function getCollectionMonitorFresh() {
  // Separate short-lived cache so new form rows appear within 60 s
  var cached = cache.get('collMonitorFresh');
  if (cached) return JSON.parse(cached);

  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  var raw   = sheet.getDataRange().getValues();
  var headers = raw[0].map(function(h) { return h.toString(); });
  var data    = raw.slice(1).filter(function(row) { return row.some(function(c) { return c !== ''; }); });
  var idx     = getColumnIndices(headers);

  var result = calculateCollectionMonitor(data, headers, idx);
  cache.put('collMonitorFresh', JSON.stringify(result), 60);
  return result;
}

function calculateCollectionMonitor(data, headers, idx) {
  var n = data.length || 1;

  // ── Field definitions ─────────────────────────────────────────────
  var allFields = [
    { name:'Patient ID',      colIdx:idx.patientId,         section:'Admin',          critical:true  },
    { name:'Admission Date',  colIdx:idx.admissionDate,     section:'Admin',          critical:true  },
    { name:'Discharge Date',  colIdx:idx.dischargeDate,     section:'Discharge',      critical:false },
    { name:'Age',             colIdx:idx.age,               section:'Demographics',   critical:true  },
    { name:'Gender',          colIdx:idx.gender,            section:'Demographics',   critical:true  },
    { name:'BMI',             colIdx:idx.bmi,               section:'Demographics',   critical:false },
    { name:'Day of Illness',  colIdx:idx.dayOfIllness,      section:'Clinical',       critical:true  },
    { name:'Fever',           colIdx:idx.fever,             section:'Symptoms',       critical:true  },
    { name:'Bleeding',        colIdx:idx.bleeding,          section:'Symptoms',       critical:true  },
    { name:'Headache',        colIdx:idx.headache,          section:'Symptoms',       critical:false },
    { name:'Hypertension',    colIdx:idx.comorbidities_ht,  section:'Comorbidities',  critical:false },
    { name:'Diabetes',        colIdx:idx.comorbidities_dm,  section:'Comorbidities',  critical:false },
    { name:'NSAIDs',          colIdx:idx.nsaids,            section:'Medications',    critical:true  },
    { name:'Platelet (Adm)',  colIdx:idx.plateletAdm,       section:'Labs',           critical:true  },
    { name:'WBC (Adm)',       colIdx:idx.wbcAdm,            section:'Labs',           critical:true  },
    { name:'ALT (Adm)',       colIdx:idx.altAdm,            section:'Labs',           critical:false },
    { name:'Platelet Nadir',  colIdx:idx.plateletLow,       section:'Follow-up Labs', critical:true  },
    { name:'WBC Nadir',       colIdx:idx.wbcLow,            section:'Follow-up Labs', critical:false },
    { name:'ALT Peak',        colIdx:idx.altHigh,           section:'Follow-up Labs', critical:false },
    { name:'NS1 Antigen',     colIdx:idx.ns1,               section:'Serology',       critical:true  },
    { name:'IgG',             colIdx:idx.igg,               section:'Serology',       critical:false },
    { name:'IgM',             colIdx:idx.igm,               section:'Serology',       critical:false },
    { name:'Final Diagnosis', colIdx:idx.finalDiagnosis,    section:'Discharge',      critical:true  },
    { name:'Outcome',         colIdx:idx.outcome,           section:'Discharge',      critical:true  },
    { name:'Hospital Days',   colIdx:idx.hospitalDays,      section:'Discharge',      critical:false },
    { name:'ICU Admission',   colIdx:idx.icuAdmit,          section:'Outcomes',       critical:false }
  ].filter(function(f) { return f.colIdx !== undefined && f.colIdx >= 0; });

  // ── Per-patient completeness ──────────────────────────────────────
  var perPatientScores = [];
  var patientGaps      = [];

  data.forEach(function(row, i) {
    var filled  = 0;
    var missing = [];
    allFields.forEach(function(f) {
      var v = row[f.colIdx];
      if (v !== '' && v !== null && v !== undefined) { filled++; }
      else { missing.push(f.name); }
    });
    var score = Math.round((filled / allFields.length) * 100);
    var pid   = idx.patientId >= 0 ? (row[idx.patientId] || ('P' + (i + 1))) : ('P' + (i + 1));
    perPatientScores.push(score);
    patientGaps.push({ patientId:pid, rowIndex:i+2, missingCount:missing.length, missingFields:missing.slice(0,8), completeness:score });
  });
  patientGaps.sort(function(a, b) { return b.missingCount - a.missingCount; });

  // ── Completeness histogram ────────────────────────────────────────
  var hist = { perfect:0, high:0, medium:0, low:0 };
  perPatientScores.forEach(function(s) {
    if (s >= 95)      hist.perfect++;
    else if (s >= 80) hist.high++;
    else if (s >= 60) hist.medium++;
    else              hist.low++;
  });

  // ── Fix today (critical fields only) ────────────────────────────
  var missCounts = {};
  allFields.forEach(function(f) {
    if (!f.critical) return;
    var cnt = data.filter(function(row) { var v = row[f.colIdx]; return v === '' || v === null || v === undefined; }).length;
    if (cnt > 0) missCounts[f.name] = cnt;
  });
  var fixToday = Object.keys(missCounts)
    .sort(function(a,b) { return missCounts[b] - missCounts[a]; })
    .slice(0, 6)
    .map(function(name) { return { field:name, count:missCounts[name], pct:((missCounts[name]/n)*100).toFixed(1) }; });

  // ── Enrollment + completeness trend (by week) ────────────────────
  var enrollByWeek = {}, compByWeek = {};
  data.forEach(function(row, i) {
    var ts = idx.timestamp >= 0 ? row[idx.timestamp] : null;
    if (!ts) return;
    var d = ts instanceof Date ? ts : new Date(ts);
    if (isNaN(d.getTime())) return;
    var startOfYear = new Date(d.getFullYear(), 0, 1);
    var wn = Math.ceil(((d - startOfYear) / 86400000 + startOfYear.getDay() + 1) / 7);
    var wk = d.getFullYear() + '-W' + (wn < 10 ? '0' : '') + wn;
    if (!enrollByWeek[wk]) { enrollByWeek[wk] = 0; compByWeek[wk] = { total:0, count:0 }; }
    enrollByWeek[wk]++;
    compByWeek[wk].total += perPatientScores[i];
    compByWeek[wk].count++;
  });
  var sortedWeeks   = Object.keys(enrollByWeek).sort();
  var enrollmentTrend = sortedWeeks.map(function(wk) {
    return { label:wk, count:enrollByWeek[wk], avgCompleteness:Math.round(compByWeek[wk].total / compByWeek[wk].count) };
  });

  // ── Days-to-completion (avg completeness by day of illness) ──────
  var dayGroups = {};
  data.forEach(function(row, i) {
    var day = Math.min(Math.max(parseInt(row[idx.dayOfIllness]) || 1, 1), 10);
    if (!dayGroups[day]) dayGroups[day] = { total:0, count:0 };
    dayGroups[day].total += perPatientScores[i];
    dayGroups[day].count++;
  });
  var daysToCompletion = [];
  for (var d = 1; d <= 10; d++) {
    if (dayGroups[d]) {
      daysToCompletion.push({ day:'Day '+d, avgCompleteness:Math.round(dayGroups[d].total/dayGroups[d].count), count:dayGroups[d].count });
    }
  }

  // ── Follow-up section tracker ─────────────────────────────────────
  function sectionPct(fIdxArr) {
    if (!fIdxArr.length) return 100;
    var filled = 0, total = data.length * fIdxArr.length;
    data.forEach(function(row) {
      fIdxArr.forEach(function(fi) { var v = row[fi]; if (v !== '' && v !== null && v !== undefined) filled++; });
    });
    return Math.round((filled / (total || 1)) * 100);
  }
  var admF   = [idx.patientId, idx.age, idx.gender, idx.dayOfIllness, idx.fever, idx.bleeding, idx.comorbidities_ht, idx.comorbidities_dm, idx.nsaids, idx.admissionDate].filter(function(i) { return i >= 0; });
  var labAdmF= [idx.plateletAdm, idx.wbcAdm, idx.altAdm, idx.astAdm, idx.tempAdm, idx.sbpAdm].filter(function(i) { return i >= 0; });
  var seroF  = [idx.ns1, idx.igg, idx.igm].filter(function(i) { return i >= 0; });
  var fuLabF = [idx.plateletLow, idx.wbcLow, idx.altHigh, idx.astHigh].filter(function(i) { return i >= 0; });
  var disF   = [idx.finalDiagnosis, idx.outcome, idx.hospitalDays, idx.dischargeDate].filter(function(i) { return i >= 0; });
  var followUpTracker = [
    { label:'Admission Data',  pct:sectionPct(admF),   icon:'🏥', desc:'Demographics, symptoms, comorbidities, NSAIDs' },
    { label:'Admission Labs',  pct:sectionPct(labAdmF),icon:'🔬', desc:'Platelet, WBC, LFT, vitals on admission' },
    { label:'Serology',        pct:sectionPct(seroF),  icon:'🧬', desc:'NS1 antigen, IgG, IgM results' },
    { label:'Follow-up Labs',  pct:sectionPct(fuLabF), icon:'📊', desc:'Platelet nadir, WBC nadir, peak LFT values' },
    { label:'Discharge Data',  pct:sectionPct(disF),   icon:'📋', desc:'Final diagnosis, outcome, length of stay' }
  ];

  var avgComp = Math.round(perPatientScores.reduce(function(a,b){return a+b;},0) / (perPatientScores.length || 1));

  return {
    totalPatients:         data.length,
    adjustedCompleteness:  avgComp,
    enrollmentTrend:       enrollmentTrend,
    daysToCompletion:      daysToCompletion,
    patientGaps:           patientGaps.slice(0, 25),
    fixToday:              fixToday,
    completenessHistogram: hist,
    followUpTracker:       followUpTracker,
    lastUpdated:           new Date().toISOString()
  };
}

function calculateLifestyleImpact(data, headers, idx) {
  var n = data.length || 1;
  var smokingIdx = headers.findIndex(function(h) { return h.toLowerCase().includes('smoking'); });
  var alcoholIdx = headers.findIndex(function(h) { return h.toLowerCase().includes('alcohol'); });
  
  var currentSmokers = 0, alcoholUsers = 0;
  var smokerBleeding = 0, smokerSevere = 0;
  var alcoholBleeding = 0, alcoholSevere = 0;
  
  data.forEach(function(row) {
    var smokingStatus = (smokingIdx >= 0 ? (row[smokingIdx] || '') : '').toString().toLowerCase();
    var isSmoker = smokingStatus.includes('current');
    var isAlcohol = alcoholIdx >= 0 && row[alcoholIdx] === 'Yes';
    var hasBleeding = row[idx.bleeding] === 'Yes';
    var severity = assessSeverity(row, idx);
    var isSevere = severity === 'Severe';
    
    if (isSmoker) {
      currentSmokers++;
      if (hasBleeding) smokerBleeding++;
      if (isSevere) smokerSevere++;
    }
    
    if (isAlcohol) {
      alcoholUsers++;
      if (hasBleeding) alcoholBleeding++;
      if (isSevere) alcoholSevere++;
    }
  });
  
  var smokingSmokers = currentSmokers || 1;
  var alcoholUsersCount = alcoholUsers || 1;
  
  return {
    currentSmokers: currentSmokers,
    alcoholUsers: alcoholUsers,
    smokingBleedingPct: ((smokerBleeding / smokingSmokers) * 100).toFixed(1),
    smokingSeverePct: ((smokerSevere / smokingSmokers) * 100).toFixed(1),
    alcoholBleedingPct: ((alcoholBleeding / alcoholUsersCount) * 100).toFixed(1),
    alcoholSeverePct: ((alcoholSevere / alcoholUsersCount) * 100).toFixed(1)
  };
}
