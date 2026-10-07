function doGet(e) {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle("Daewoong Global Recruitment Portal")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 1. Authenticate User & Baca Group/Category dari Kolom D tab Users
function authenticateUser(username, password, role) {
  if (role === 'Pelamar') {
    return { success: true, role: 'Pelamar' };
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let userSheet = ss.getSheetByName("Users");
  
  // Buat sheet Users jika belum ada
  if (!userSheet) {
    userSheet = ss.insertSheet("Users");
    userSheet.appendRow(["Nama Evaluator", "Password", "Role", "Assigned Categories"]);
    userSheet.appendRow(["유지민", "eval123", "Evaluator", "Quality A, Quality B, Research"]);
    userSheet.appendRow(["이승하", "eval123", "Evaluator", "Production, Engineering"]);
    userSheet.appendRow(["합세형", "eval123", "Evaluator", "Quality A, Quality"]);
    userSheet.appendRow(["admin", "admin123", "Admin", "ALL"]);
  }
  
  const data = userSheet.getDataRange().getValues();
  const inputUser = username ? username.toString().trim().toLowerCase() : "";
  const inputPass = password ? password.toString().trim() : "";
  
  for (let i = 1; i < data.length; i++) {
    const realName = data[i][0] ? data[i][0].toString().trim() : "";
    const uPass = data[i][1] ? data[i][1].toString().trim() : "";
    const uRole = data[i][2] ? data[i][2].toString().trim() : "Evaluator";
    const uCategories = data[i][3] ? data[i][3].toString().trim() : "";
    
    if (realName.toLowerCase() === inputUser && uPass === inputPass) {
      return { 
        success: true, 
        username: realName,
        role: uRole.toLowerCase() === 'admin' ? 'Admin' : 'Evaluator',
        assignedCategories: uCategories
      };
    }
  }
  
  return { success: false, message: "Invalid Name, Password, or Role selection!" };
}

// 2. Submit Application Form (Pelamar)
function processApplication(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Kandidat_Master");
    
    if (!sheet) {
      sheet = ss.insertSheet("Kandidat_Master");
      sheet.appendRow([
        "ID Kandidat", "Nama", "Email", "Category", "Phone Number", "Birthday",
        "TOP 5 University", "University Name", "Degree", "Pharmacy License", "Major",
        "Nationality", "English Proficiency", "Korean Proficiency", "Experience Status",
        "Total Work Experienced", "Religion", "Address", "IT Experience", "Link CV",
        "Link Application Form", "Form", "Link Video", "Doc Screening (Admin)", 
        "Aptitude Test (Admin)", "Video Eval (Evaluator)", "Interview Eval (Evaluator)", 
        "Final Result (Admin)", "Evaluator Notes", "Video Evaluators", "Interview Evaluators", "Timestamp"
      ]);
    }
    
    const timestamp = new Date();
    const candidateId = "CAND-" + Utilities.formatDate(timestamp, "GMT+7", "yyyyMMdd") + "-" + Math.floor(1000 + Math.random() * 9000);
    
    sheet.appendRow([
      candidateId, formData.nama, formData.email, formData.category, formData.phone,
      formData.birthday, formData.university, formData.schoolName, formData.degree,
      formData.pharmacyLicense, formData.major, formData.nationality, formData.englishProficiency,
      formData.koreanProficiency, formData.experienceStatus, formData.workExperience,
      formData.religion, formData.address, formData.itExperience, formData.cvUrl || "",
      formData.appFormUrl || "", "", formData.videoUrl || "", 
      "Pending", "Pending", "Pending", "Pending", "Pending", "", "", "", timestamp
    ]);
    
    return { success: true, candidateId: candidateId };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// 3. Get Candidates Data (Auto-Filtered berdasarkan Category)
function getCandidatesData(username, userRole, assignedCategories) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Kandidat_Master");
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (!data || data.length <= 1) return [];
    
    const candidates = [];
    
    const categoryList = assignedCategories 
      ? assignedCategories.split(',').map(function(c) { return c.trim().toLowerCase(); }) 
      : [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row || (!row[0] && !row[1])) continue;
      
      const candidateCategory = row[3] ? row[3].toString().trim().toLowerCase() : "";
      const videoEvalsStr = row[29] ? row[29].toString().trim() : "";
      const interviewEvalsStr = row[30] ? row[30].toString().trim() : "";
      
      if (userRole === 'Evaluator') {
        const hasCategoryAccess = categoryList.some(function(cat) { 
          return candidateCategory.indexOf(cat) !== -1 || cat.indexOf(candidateCategory) !== -1; 
        });
        const isDirectlyAssigned = videoEvalsStr.toLowerCase().indexOf(username.toLowerCase()) !== -1 || 
                                   interviewEvalsStr.toLowerCase().indexOf(username.toLowerCase()) !== -1;

        if (!hasCategoryAccess && !isDirectlyAssigned && categoryList.length > 0) {
          continue;
        }
      }

      candidates.push({
        rowIndex: i + 1,
        candidateId: row[0] ? row[0].toString() : ("CAND-" + i),
        nama: row[1] ? row[1].toString() : "-",
        email: row[2] ? row[2].toString() : "-",
        category: row[3] ? row[3].toString() : "-",
        phone: row[4] ? row[4].toString() : "-",
        birthday: row[5] ? row[5].toString() : "-",
        university: row[6] ? row[6].toString() : "-",
        schoolName: row[7] ? row[7].toString() : "-",
        degree: row[8] ? row[8].toString() : "-",
        pharmacyLicense: row[9] ? row[9].toString() : "-",
        major: row[10] ? row[10].toString() : "-",
        nationality: row[11] ? row[11].toString() : "-",
        englishProficiency: row[12] ? row[12].toString() : "-",
        koreanProficiency: row[13] ? row[13].toString() : "-",
        experienceStatus: row[14] ? row[14].toString() : "-",
        workExperience: row[15] ? row[15].toString() : "-",
        religion: row[16] ? row[16].toString() : "-",
        address: row[17] ? row[17].toString() : "-",
        itExperience: row[18] ? row[18].toString() : "-",
        cvUrl: row[19] ? row[19].toString() : "",
        appFormUrl: row[20] ? row[20].toString() : "",
        videoUrl: (row[22] && row[22].toString().indexOf("http") === 0) ? row[22].toString() : ((row[21] && row[21].toString().indexOf("http") === 0) ? row[21].toString() : ""),
        docScreening: row[23] ? row[23].toString() : "Pending",
        aptitudeTest: row[24] ? row[24].toString() : "Pending",
        videoEval: row[25] ? row[25].toString() : "Pending",
        interviewEval: row[26] ? row[26].toString() : "Pending",
        finalResult: row[27] ? row[27].toString() : "Pending",
        evalNotes: row[28] ? row[28].toString() : "",
        videoEvaluators: videoEvalsStr,
        interviewEvaluators: interviewEvalsStr
      });
    }
    
    return candidates;
  } catch (err) {
    Logger.log("Error in getCandidatesData: " + err.toString());
    return [];
  }
}

// 4. Update Evaluasi per Stage (Admin & Evaluator)
function updateEvaluationStage(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Kandidat_Master");
    const rowIndex = parseInt(payload.rowIndex);
    
    if (payload.docScreening !== undefined) sheet.getRange(rowIndex, 24).setValue(payload.docScreening);
    if (payload.aptitudeTest !== undefined) sheet.getRange(rowIndex, 25).setValue(payload.aptitudeTest);
    if (payload.videoEval !== undefined) sheet.getRange(rowIndex, 26).setValue(payload.videoEval);
    if (payload.interviewEval !== undefined) sheet.getRange(rowIndex, 27).setValue(payload.interviewEval);
    if (payload.finalResult !== undefined) sheet.getRange(rowIndex, 28).setValue(payload.finalResult);
    if (payload.notes !== undefined) sheet.getRange(rowIndex, 29).setValue(payload.notes);
    
    return { success: true };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// 5. Submit Multi-Evaluator Log ke Tab Evaluations
function submitEvaluatorScore(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let evalSheet = ss.getSheetByName("Evaluations");
    
    if (!evalSheet) {
      evalSheet = ss.insertSheet("Evaluations");
      evalSheet.appendRow(["Candidate ID", "Evaluator Username", "Stage", "Decision", "Notes", "Timestamp"]);
    }
    
    const timestamp = new Date();
    evalSheet.appendRow([
      payload.candidateId,
      payload.evaluatorUsername,
      payload.stage,
      payload.decision,
      payload.notes || "",
      timestamp
    ]);
    
    const masterSheet = ss.getSheetByName("Kandidat_Master");
    const rowIndex = parseInt(payload.rowIndex);
    
    if (payload.stage === 'Video Evaluation') {
      masterSheet.getRange(rowIndex, 26).setValue(payload.decision);
    } else if (payload.stage === 'Interview Evaluation') {
      masterSheet.getRange(rowIndex, 27).setValue(payload.decision);
    }
    
    return { success: true };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// 6. Ambil History Multi-Evaluator per Kandidat
function getCandidateEvaluationHistory(candidateId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const evalSheet = ss.getSheetByName("Evaluations");
    if (!evalSheet) return [];
    
    const data = evalSheet.getDataRange().getValues();
    const history = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === candidateId.trim()) {
        history.push({
          evaluator: data[i][1],
          stage: data[i][2],
          decision: data[i][3],
          notes: data[i][4],
          timestamp: Utilities.formatDate(new Date(data[i][5]), "GMT+7", "yyyy-MM-dd HH:mm")
        });
      }
    }
    return history;
  } catch (err) {
    return [];
  }
}
