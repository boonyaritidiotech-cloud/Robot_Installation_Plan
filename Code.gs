var CALENDAR_ID = "3vsvb0a1nblh7odvi0qecsb0a0@group.calendar.google.com";

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Robot Operation System V7.7')
    .addMetaTag('viewport','width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  try { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }
  catch(e) { return ""; }
}

// ============================================================
// AUTH
// ============================================================
function checkLogin(u, p) {
  try {
    var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users').getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === u.trim() && data[i][1].toString().trim() === p.trim()) {
        var role = data[i][4];
        return { success:true, name:data[i][2], role:role, permissions:getRolePermissions(role) };
      }
    }
    return { success:false, msg:"Username หรือ Password ไม่ถูกต้อง" };
  } catch(e) { return { success:false, msg:"ไม่พบฐานข้อมูลผู้ใช้: "+e.toString() }; }
}

// ============================================================
// ROLES
// ============================================================
function getRolePermissions(role) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Roles');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === role.toString().trim()) {
        var perms = {};
        for (var j = 1; j < headers.length; j++)
          perms[headers[j]] = data[i][j] === true || data[i][j] === 'TRUE';
        return perms;
      }
    }
    return {};
  } catch(e) { return {}; }
}

function getRoleList() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Roles');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    return data.slice(1).map(function(row) {
      var perms = {};
      for (var j = 1; j < headers.length; j++)
        perms[headers[j]] = row[j] === true || row[j] === 'TRUE';
      return { role:row[0], permissions:perms };
    });
  } catch(e) { return []; }
}

function createRole(roleName, permissions) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Roles');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    for (var i = 1; i < data.length; i++)
      if (data[i][0].toString().trim().toLowerCase() === roleName.trim().toLowerCase())
        return { success:false, msg:"Role นี้มีอยู่แล้ว" };
    var newRow = [roleName];
    for (var j = 1; j < headers.length; j++) newRow.push(permissions[headers[j]] === true);
    sheet.appendRow(newRow);
    return { success:true, msg:"สร้าง Role สำเร็จ!" };
  } catch(e) { return { success:false, msg:e.toString() }; }
}

function updateRole(roleName, permissions) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Roles');
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === roleName.trim()) {
        for (var j = 1; j < headers.length; j++)
          sheet.getRange(i+1,j+1).setValue(permissions[headers[j]] === true);
        return { success:true, msg:"อัปเดต Role สำเร็จ!" };
      }
    }
    return { success:false, msg:"ไม่พบ Role นี้" };
  } catch(e) { return { success:false, msg:e.toString() }; }
}

function deleteRole(roleName) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Roles');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === roleName.trim()) {
        sheet.deleteRow(i+1);
        return { success:true, msg:"ลบ Role สำเร็จ!" };
      }
    }
    return { success:false, msg:"ไม่พบ Role นี้" };
  } catch(e) { return { success:false, msg:e.toString() }; }
}

// ============================================================
// INITIAL DATA
// ============================================================
function getInitialData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var robots = ss.getSheetByName('Robots').getDataRange().getValues().slice(1)
      .map(function(r){ return { brand:r[0], model:r[1] }; });
    var custMap = {};
    ss.getSheetByName('Customers').getDataRange().getValues().slice(1).forEach(function(c) {
      var company = c[0] ? c[0].toString().trim() : "";
      var branch  = c[1] ? c[1].toString().trim() : "";
      if (!company) return;
      if (!custMap[company]) custMap[company] = [];
      if (branch && custMap[company].indexOf(branch) === -1) custMap[company].push(branch);
    });
    var customers = Object.keys(custMap).map(function(co){ return { company:co, branches:custMap[co] }; });
    var staff = {};
    ss.getSheetByName('Users').getDataRange().getValues().slice(1).forEach(function(u) {
      var dept = u[5] ? u[5].toString().trim() : "Other";
      var name = u[2] ? u[2].toString().trim() : "";
      if (!name) return;
      if (!staff[dept]) staff[dept] = [];
      staff[dept].push({ name:name });
    });
    var roles = getRoleList().map(function(r){ return r.role; });
    return { robots:robots, customers:customers, staff:staff, roles:roles };
  } catch(e) { return { robots:[], customers:[], staff:{}, roles:[] }; }
}

function buildEmailMap() {
  var map = {};
  try {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users')
      .getDataRange().getValues().slice(1).forEach(function(row) {
        var name  = row[2] ? row[2].toString().trim() : "";
        var email = row[3] ? row[3].toString().trim() : "";
        if (name && email) map[name] = email;
      });
  } catch(e) { Logger.log("buildEmailMap Error: "+e); }
  return map;
}

// ============================================================
// HELPERS
// ============================================================
function toDateStr(val) {
  if (!val) return "";
  if (val instanceof Date) return Utilities.formatDate(val, "GMT+7", "yyyy-MM-dd");
  var s = val.toString().trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  try { return Utilities.formatDate(new Date(s), "GMT+7", "yyyy-MM-dd"); } catch(e) { return s; }
}

function toTimeStr(val) {
  if (!val) return "";
  if (val instanceof Date) {
    var h = val.getHours().toString().padStart(2,'0');
    var m = val.getMinutes().toString().padStart(2,'0');
    return h + ":" + m;
  }
  var s = val.toString().trim();
  if (s.length > 5) s = s.substring(0, 5);
  if (/^\d:\d{2}$/.test(s)) s = "0" + s;
  return s;
}

// ✅ แก้ปัญหางานข้ามคืน — auto เพิ่ม end date +1 ถ้า eTime < sTime
function fixOvernightJob(startStr, endStr, sTimeStr, eTimeStr) {
  var sd = new Date(startStr + "T" + sTimeStr + ":00");
  var ed = new Date(endStr   + "T" + eTimeStr + ":00");
  if (ed <= sd) {
    var nextDay = new Date(endStr);
    nextDay.setDate(nextDay.getDate() + 1);
    endStr = Utilities.formatDate(nextDay, "GMT+7", "yyyy-MM-dd");
    ed = new Date(endStr + "T" + eTimeStr + ":00");
    Logger.log("Overnight job — end adjusted to: " + endStr);
  }
  return { startStr:startStr, endStr:endStr, sd:sd, ed:ed };
}

function mapJobRow(row) {
  return {
    id:         row[0] ? row[0].toString() : "",
    start:      row[1] ? toDateStr(row[1]) : "",
    end:        row[2] ? toDateStr(row[2]) : "",
    date:       row[1] ? Utilities.formatDate(new Date(row[1]),"GMT+7","dd/MM/yyyy") : "",
    endDate:    row[2] ? Utilities.formatDate(new Date(row[2]),"GMT+7","dd/MM/yyyy") : "",
    sTime:      toTimeStr(row[3]),
    eTime:      toTimeStr(row[4]),
    cust:       (row[5]||"").toString(),
    branch:     (row[6]||"").toString(),
    jobType:    (row[7]||"").toString(),
    robots:     (row[8]||"").toString(),
    staff:      (row[10]||"").toString(),
    remark:     (row[11]||"").toString(),
    eventId:    (row[12]||"").toString(),
    status:     (row[13]||"งานใหม่").toString(),
    updatedBy:  (row[14]||"").toString(),
    updateTime: (row[15]||"").toString()
  };
}

function rowToRequest(row) {
  return {
    reqId:      row[0]  ? row[0].toString()  : "",
    date:       row[1]  ? row[1].toString()  : "",
    cust:       row[2]  ? row[2].toString()  : "",
    branch:     row[3]  ? row[3].toString()  : "",
    location:   row[4]  ? row[4].toString()  : "",
    start:      toDateStr(row[5]),
    end:        toDateStr(row[6]),
    sTime:      toTimeStr(row[7]),
    eTime:      toTimeStr(row[8]),
    robots:     row[9]  ? row[9].toString()  : "",
    staff:      row[10] ? row[10].toString() : "",
    remark:     row[11] ? row[11].toString() : "",
    requestBy:  row[12] ? row[12].toString() : "",
    status:     row[13] ? row[13].toString() : "รอดำเนินการ",
    approveBy:  row[14] ? row[14].toString() : "",
    reason:     row[15] ? row[15].toString() : "",
    updateTime: row[16] ? row[16].toString() : "",
    jobType:    row[17] ? row[17].toString() : ""
  };
}

// ============================================================
// CONFLICT CHECK
// ============================================================
function checkStaffConflict(start, end, sTime, eTime, staffList, excludeJobId) {
  try {
    // ✅ สร้าง map: fullname → username และ username → fullname
    var nameMap = {};
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users')
      .getDataRange().getValues().slice(1).forEach(function(u) {
        var username = u[0] ? u[0].toString().trim() : "";
        var fullname = u[2] ? u[2].toString().trim() : "";
        if (username && fullname) {
          nameMap[fullname.toLowerCase()] = username.toLowerCase();
          nameMap[username.toLowerCase()] = username.toLowerCase();
        }
      });

    var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data').getDataRange().getValues();
    var conflicts = [];
    var reqStart = new Date(start+"T"+sTime+":00");
    var reqEnd   = new Date(end  +"T"+eTime+":00");

    data.slice(1).forEach(function(row) {
      if (excludeJobId && row[0].toString() === excludeJobId.toString()) return;
      if (!row[1] || !row[2]) return;
      var rs = toDateStr(row[1]);
      var re = toDateStr(row[2]);
      var st = toTimeStr(row[3]) || "00:00";
      var et = toTimeStr(row[4]) || "23:59";
      var rowStart = new Date(rs+"T"+st+":00");
      var rowEnd   = new Date(re+"T"+et+":00");
      if (reqStart >= rowEnd || reqEnd <= rowStart) return;

      var assigned = row[10] ? row[10].toString().split(",")
        .map(function(s){ return s.trim().toLowerCase(); }) : [];

      staffList.forEach(function(name) {
        // ✅ เช็คด้วย username ที่ normalize แล้ว
        var nameKey = nameMap[name.trim().toLowerCase()] || name.trim().toLowerCase();
        var found = assigned.some(function(a) {
          var aKey = nameMap[a] || a;
          return aKey === nameKey;
        });
        if (found) conflicts.push({
          name: name,
          jobId: row[0],
          date: Utilities.formatDate(new Date(row[1]),"GMT+7","dd/MM/yyyy")
        });
      });
    });
    return conflicts;
  } catch(e) { Logger.log("Conflict Error: "+e); return []; }
}

// ============================================================
// JOBS
// ============================================================
function saveJob(obj) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Data');
    var jobId = "J-" + new Date().getTime();

    var startStr = toDateStr(obj.start);
    var endStr   = toDateStr(obj.end);
    var sTimeStr = toTimeStr(obj.sTime) || "09:00";
    var eTimeStr = toTimeStr(obj.eTime) || "17:00";
    var location = (obj.location||"") + (obj.branch ? " ("+obj.branch+")" : "");
    var staffArr = Array.isArray(obj.staff) ? obj.staff : (obj.staff||"").split(",").map(function(s){ return s.trim(); }).filter(Boolean);
    var staffStr = staffArr.join(", ");

    // ✅ fix overnight
    var fixed = fixOvernightJob(startStr, endStr, sTimeStr, eTimeStr);
    startStr = fixed.startStr;
    endStr   = fixed.endStr;

    // Conflict check
    if (staffArr.length > 0) {
      var conflicts = checkStaffConflict(startStr, endStr, sTimeStr, eTimeStr, staffArr, null);
      if (conflicts.length > 0) {
        var msg = conflicts.map(function(c){ return c.name+" (Job: "+c.jobId+" วันที่ "+c.date+")"; }).join(", ");
        return { success:false, msg:"⚠️ พบงานชน: "+msg };
      }
    }

    // Calendar
    var eventId = "";
    try {
      var cal = CalendarApp.getCalendarById(CALENDAR_ID);
      var ev  = cal.createEvent(
        "["+obj.cust+(obj.branch?" - "+obj.branch:"")+"] "+obj.robots,
        fixed.sd, fixed.ed,
        { description:
            "📍 "+location+
            "\n👥 Staff: "+staffStr+
            "\n📋 "+(obj.jobType||"-")+
            "\n📝 "+(obj.remark||"-")+
            "\n🔖 Job ID: "+jobId,
          location: location }
      );
      eventId = ev.getId();
      Logger.log("saveJob Calendar created: "+eventId);
    } catch(e) { Logger.log("Cal Error: "+e.toString()); }

    // Sheet
    sheet.appendRow([
      jobId, startStr, endStr, sTimeStr, eTimeStr,
      obj.cust, obj.branch||"", obj.jobType||"Service", obj.robots, "",
      staffStr, obj.remark||"", eventId, "งานใหม่", "", ""
    ]);

    // Email
    try {
      var emailMap = buildEmailMap();
      staffArr.forEach(function(name) {
        var email = emailMap[name.trim()]; if(!email) return;
        MailApp.sendEmail(email, "📋 งานใหม่: "+obj.cust,
          "สวัสดีคุณ "+name+",\n\nคุณได้รับมอบหมายงานใหม่\n\n━━━━━━━━━━━━━━━━━━━━\n"+
          "🔖 Job ID   : "+jobId+"\n🏢 ลูกค้า   : "+obj.cust+(obj.branch?" ("+obj.branch+")":"")+"\n"+
          "📋 ประเภท   : "+(obj.jobType||"-")+"\n📍 สถานที่  : "+location+"\n"+
          "📅 วันที่    : "+startStr+" ถึง "+endStr+"\n"+
          "⏰ เวลา     : "+sTimeStr+" - "+eTimeStr+"\n"+
          "🤖 หุ่นยนต์ : "+obj.robots+"\n"+
          "📝 หมายเหตุ : "+(obj.remark||"-")+"\n"+
          "━━━━━━━━━━━━━━━━━━━━\n\nRobot Operation System"
        );
      });
    } catch(e) { Logger.log("Email Error: "+e.toString()); }

    return { success:true, msg:"บันทึกแผนงานสำเร็จ! ✅", jobId:jobId };
  } catch(e) { return { success:false, msg:"saveJob Error: "+e.toString() }; }
}

function getJobDetail(jobId) {
  try {
    var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data').getDataRange().getValues();
    for (var i = 1; i < data.length; i++)
      if (data[i][0].toString() === jobId) return mapJobRow(data[i]);
    return null;
  } catch(e) { return null; }
}

function updateJob(jobId, obj) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
    var data  = sheet.getDataRange().getValues();

    var startStr = toDateStr(obj.start);
    var endStr   = toDateStr(obj.end);
    var sTimeStr = toTimeStr(obj.sTime) || "09:00";
    var eTimeStr = toTimeStr(obj.eTime) || "17:00";
    var fixed    = fixOvernightJob(startStr, endStr, sTimeStr, eTimeStr);
    startStr = fixed.startStr;
    endStr   = fixed.endStr;

    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString() !== jobId) continue;

      var staffList = Array.isArray(obj.staff)
        ? obj.staff
        : obj.staff.split(",").map(function(s){ return s.trim(); }).filter(Boolean);

      if (staffList.length > 0) {
        var conflicts = checkStaffConflict(startStr, endStr, sTimeStr, eTimeStr, staffList, jobId);
        if (conflicts.length > 0) {
          var msg = conflicts.map(function(c){ return c.name+" (Job: "+c.jobId+" วันที่ "+c.date+")"; }).join(", ");
          return { success:false, msg:"⚠️ พบงานชน: "+msg };
        }
      }

      var now        = Utilities.formatDate(new Date(),"GMT+7","dd/MM/yyyy HH:mm:ss");
      var location   = (obj.location||"")+(obj.branch?" ("+obj.branch+")":"");
      var oldEventId = data[i][12] ? data[i][12].toString().trim() : "";
      var newEventId = "";

      // ✅ ลบ event เก่า + สร้างใหม่
      try {
        var cal = CalendarApp.getCalendarById(CALENDAR_ID);
        if (oldEventId) {
          try {
            var oldEv = cal.getEventById(oldEventId);
            if (oldEv) { oldEv.deleteEvent(); Logger.log("Deleted: "+oldEventId); }
          } catch(e) { Logger.log("Delete error: "+e); }
        }
        var newEv = cal.createEvent(
          "["+obj.cust+(obj.branch?" - "+obj.branch:"")+"] "+obj.robots,
          fixed.sd, fixed.ed,
          { description:
              "📍 "+location+
              "\n👥 Staff: "+staffList.join(", ")+
              "\n📋 "+(obj.jobType||"-")+
              "\n📝 "+(obj.remark||"-")+
              "\n🔖 Job ID: "+jobId,
            location: location }
        );
        newEventId = newEv.getId();
        Logger.log("Created new event: "+newEventId);
      } catch(e) { Logger.log("Cal Error: "+e); }

      // อัปเดต Sheet
      sheet.getRange(i+1,2).setValue(startStr);
      sheet.getRange(i+1,3).setValue(endStr);
      sheet.getRange(i+1,4).setValue(sTimeStr);
      sheet.getRange(i+1,5).setValue(eTimeStr);
      sheet.getRange(i+1,6).setValue(obj.cust);
      sheet.getRange(i+1,7).setValue(obj.branch||"");
      sheet.getRange(i+1,8).setValue(obj.jobType||data[i][7]||"Service");
      sheet.getRange(i+1,9).setValue(obj.robots);
      sheet.getRange(i+1,11).setValue(staffList.join(", "));
      sheet.getRange(i+1,12).setValue(obj.remark||"");
      sheet.getRange(i+1,13).setValue(newEventId);
      sheet.getRange(i+1,15).setValue(obj.updatedBy||"");
      sheet.getRange(i+1,16).setValue(now);

      return { success:true, msg:"แก้ไขแผนงานสำเร็จ! ✅" };
    }
    return { success:false, msg:"ไม่พบรหัสงาน" };
  } catch(e) { return { success:false, msg:"updateJob Error: "+e.toString() }; }
}

function getFilteredJobList() {
  try {
    var data  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data').getDataRange().getValues();
    if (data.length <= 1) return [];
    var now   = new Date();
    var today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    return data.slice(1)
      .filter(function(row) {
        if (!row[2]) return false;
        var endDate = new Date(row[2]);
        endDate = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());
        if (endDate.getTime() === today.getTime()) {
          var eTime = toTimeStr(row[4]) || "23:59";
          var parts = eTime.split(":");
          var endDt = new Date(endDate);
          endDt.setHours(parseInt(parts[0]||23), parseInt(parts[1]||59), 0);
          return endDt >= now;
        }
        return endDate >= today;
      })
      .sort(function(a,b){ return new Date(a[1]) - new Date(b[1]); })
      .map(function(row) {
        return {
          id:       row[0] ? row[0].toString() : "",
          customer: row[5] ? row[5].toString() : "",
          branch:   row[6] ? row[6].toString() : "",
          jobType:  row[7] ? row[7].toString() : "",
          date:     row[1] ? Utilities.formatDate(new Date(row[1]),"GMT+7","dd/MM/yyyy") : "",
          endDate:  row[2] ? Utilities.formatDate(new Date(row[2]),"GMT+7","dd/MM/yyyy") : "",
          sTime:    toTimeStr(row[3]),
          eTime:    toTimeStr(row[4]),
          status:   row[13] ? row[13].toString() : "งานใหม่",
          staff:    row[10] ? row[10].toString() : "",
          robots:   row[8]  ? row[8].toString()  : ""
        };
      });
  } catch(e) { Logger.log("getFilteredJobList Error: "+e); return []; }
}

function removeJobFromServer(jobId) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
    var data  = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString() !== jobId) continue;
      var eventId  = data[i][12] ? data[i][12].toString().trim() : "";
      var staffStr = data[i][10] ? data[i][10].toString() : "";
      var cust     = data[i][5]  ? data[i][5].toString()  : "";
      var branch   = data[i][6]  ? data[i][6].toString()  : "";
      var startD   = data[i][1]  ? Utilities.formatDate(new Date(data[i][1]),"GMT+7","dd/MM/yyyy") : "";
      var robots   = data[i][8]  ? data[i][8].toString()  : "";
      if (eventId) {
        try {
          var ev = CalendarApp.getCalendarById(CALENDAR_ID).getEventById(eventId);
          if (ev) ev.deleteEvent();
        } catch(e) { Logger.log("Cal Delete Error: "+e); }
      }
      try {
        var emailMap = buildEmailMap();
        staffStr.split(",").forEach(function(name) {
          name = name.trim();
          var email = emailMap[name]; if(!email) return;
          MailApp.sendEmail(email, "❌ ยกเลิกงาน: "+cust,
            "สวัสดีคุณ "+name+",\n\nงานถูกยกเลิกแล้ว\n\n━━━━━━━━━━━━━━━━━━━━\n"+
            "🔖 Job ID : "+jobId+"\n🏢 ลูกค้า : "+cust+(branch?" ("+branch+")":"")+"\n"+
            "📅 วันที่  : "+startD+"\n🤖 หุ่นยนต์: "+robots+"\n"+
            "━━━━━━━━━━━━━━━━━━━━\n\nRobot Operation System"
          );
        });
      } catch(e) { Logger.log("Email Cancel Error: "+e); }
      sheet.deleteRow(i+1);
      return { success:true, msg:"ลบงานสำเร็จ! ✅" };
    }
    return { success:false, msg:"ไม่พบรหัสงาน" };
  } catch(e) { return { success:false, msg:"removeJob Error: "+e.toString() }; }
}

function getAllJobs() {
  try {
    var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data').getDataRange().getValues();
    if (data.length <= 1) return [];
    return data.slice(1).map(function(row){ return mapJobRow(row); })
      .sort(function(a,b){
        var o={"งานใหม่":0,"กำลังดำเนินการ":1,"เสร็จสิ้น":2};
        return (o[a.status]||0)-(o[b.status]||0);
      });
  } catch(e) { return []; }
}

function getUserJobs(userName) {
  try {
    var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data').getDataRange().getValues();
    if (data.length <= 1) return [];
    var nameLower = userName.trim().toLowerCase();
    return data.slice(1)
      .filter(function(row) {
        if (!row[10]) return false;
        return row[10].toString().split(",")
          .map(function(s){ return s.trim().toLowerCase(); })
          .indexOf(nameLower) > -1;
      })
      .map(function(row){ return mapJobRow(row); })
      .sort(function(a,b){
        var o={"งานใหม่":0,"กำลังดำเนินการ":1,"เสร็จสิ้น":2};
        return (o[a.status]||0)-(o[b.status]||0);
      });
  } catch(e) { Logger.log("getUserJobs Error: "+e); return []; }
}

function updateJobStatus(jobId, userName, status) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
    var data  = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString() === jobId) {
        var now = Utilities.formatDate(new Date(),"GMT+7","dd/MM/yyyy HH:mm:ss");
        sheet.getRange(i+1,14).setValue(status);
        sheet.getRange(i+1,15).setValue(userName);
        sheet.getRange(i+1,16).setValue(now);
        return { success:true, msg:"อัปเดตสถานะเป็น \""+status+"\" สำเร็จ!" };
      }
    }
    return { success:false, msg:"ไม่พบรหัสงาน" };
  } catch(e) { return { success:false, msg:e.toString() }; }
}

// ============================================================
// AVAILABLE STAFF
// ============================================================
function getAvailableStaffByDept(start, end, sTime, eTime) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var usernameToFullname = {};
    var fullnameToUsername = {};
    var allStaff = {};

    ss.getSheetByName('Users').getDataRange().getValues().slice(1).forEach(function(u) {
      var username = u[0] ? u[0].toString().trim() : "";
      var fullname = u[2] ? u[2].toString().trim() : "";
      var dept     = u[5] ? u[5].toString().trim() : "Other";
      if (!fullname) return;
      usernameToFullname[username.toLowerCase()] = fullname;
      fullnameToUsername[fullname.toLowerCase()] = username.toLowerCase();
      if (!allStaff[dept]) allStaff[dept] = [];
      allStaff[dept].push(fullname);
    });

    var startStr = toDateStr(start);
    var endStr   = toDateStr(end);
    var sTimeStr = toTimeStr(sTime) || "00:00";
    var eTimeStr = toTimeStr(eTime) || "23:59";
    var fixed    = fixOvernightJob(startStr, endStr, sTimeStr, eTimeStr);
    var reqStart = fixed.sd;
    var reqEnd   = fixed.ed;

    // ✅ busyMap เก็บ key=name, value=array of {date, cust}
    var busyMap = {};
    ss.getSheetByName('Data').getDataRange().getValues().slice(1).forEach(function(row) {
      if (!row[1]||!row[2]) return;
      var rs = toDateStr(row[1]);
      var re = toDateStr(row[2]);
      var st = toTimeStr(row[3]) || "00:00";
      var et = toTimeStr(row[4]) || "23:59";
      if (reqStart >= new Date(re+"T"+et+":00") || reqEnd <= new Date(rs+"T"+st+":00")) return;
      
      var jobDate = rs === re ? rs : rs+" → "+re;
      var cust    = (row[5]||"").toString();
      
      if (!row[10]) return;
      row[10].toString().split(",").forEach(function(s) {
        var key = s.trim().toLowerCase();
        // normalize เป็น fullname lowercase
        var fnKey = usernameToFullname[key] ? usernameToFullname[key].toLowerCase() : key;
        if (!busyMap[fnKey]) busyMap[fnKey] = [];
        busyMap[fnKey].push({ date:jobDate, cust:cust });
        // mark username ด้วย
        var unKey = fullnameToUsername[key] || key;
        if (!busyMap[unKey]) busyMap[unKey] = [];
        busyMap[unKey].push({ date:jobDate, cust:cust });
      });
    });

    var result = {};
    Object.keys(allStaff).forEach(function(dept) {
      result[dept] = allStaff[dept].map(function(fullname) {
        var fnKey    = fullname.toLowerCase();
        var unKey    = fullnameToUsername[fnKey] || fnKey;
        var busyJobs = busyMap[fnKey] || busyMap[unKey] || [];
        return {
          name:      fullname,
          available: busyJobs.length === 0,
          busyJobs:  busyJobs  // ✅ ส่ง busyJobs ไปด้วย
        };
      });
    });
    return result;
  } catch(e) { Logger.log("AvailStaff Error: "+e); return {}; }
}

// ============================================================
// REQUESTS
// ============================================================
function getRequestsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Requests');
  if (!sheet) {
    sheet = ss.insertSheet('Requests');
    sheet.appendRow(['Req_ID','Date','Customer','Branch','Location','Start','End',
      'sTime','eTime','Robots','Staff','Remark','RequestBy','Status',
      'ApproveBy','Reason','UpdateTime','JobType']);
    // ✅ บังคับ column sTime(H) และ eTime(I) เป็น Plain Text
    sheet.getRange('H:H').setNumberFormat('@');
    sheet.getRange('I:I').setNumberFormat('@');
  }
  return sheet;
}

function submitRequest(obj) {
  try {
    var sheet = getRequestsSheet();
    var reqId = "R-"+new Date().getTime();
    var now   = Utilities.formatDate(new Date(),"GMT+7","dd/MM/yyyy HH:mm:ss");
    sheet.appendRow([reqId, now, obj.cust, obj.branch, obj.location,
    obj.start, obj.end, "'" + (obj.sTime||"09:00"), "'" + (obj.eTime||"17:00"), obj.robots,
    obj.staff||"", obj.remark||"", obj.requestBy,
    "รอดำเนินการ","","", now, obj.jobType||""]);
    try {
      var users = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users').getDataRange().getValues().slice(1);
      users.forEach(function(u) {
        var perms = getRolePermissions(u[4]?u[4].toString().trim():"");
        if (!perms['approve']) return;
        var email = u[3]?u[3].toString().trim():""; if(!email) return;
        MailApp.sendEmail(email,"📬 คำขอลงแผนงานใหม่ จาก "+obj.requestBy,
          "มีคำขอลงแผนงานใหม่รอการอนุมัติ\n\n━━━━━━━━━━━━━━━━━━━━\n"+
          "🔖 Req ID    : "+reqId+"\n👤 ผู้ขอ     : "+obj.requestBy+"\n"+
          "📋 ประเภทงาน: "+(obj.jobType||"-")+"\n🏢 ลูกค้า   : "+obj.cust+"\n"+
          "📅 วันที่    : "+obj.start+" ถึง "+obj.end+"\n"+
          "🤖 หุ่นยนต์  : "+obj.robots+"\n━━━━━━━━━━━━━━━━━━━━\n\nRobot Operation System"
        );
      });
    } catch(e) { Logger.log("Email Error: "+e); }
    return { success:true, msg:"ส่งคำขอสำเร็จ! รอ Management อนุมัติ ✅" };
  } catch(e) { return { success:false, msg:"submitRequest Error: "+e.toString() }; }
}

function getMyRequests(userName) {
  try {
    var sheet = getRequestsSheet();
    var data  = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var results = [];
    for (var i = 1; i < data.length; i++)
      if (data[i][12] && data[i][12].toString().trim() === userName.trim())
        results.push(rowToRequest(data[i]));
    results.sort(function(a,b){
      var o={"ไม่อนุมัติ":0,"รอดำเนินการ":1,"อนุมัติ":2};
      return (o[a.status]!==undefined?o[a.status]:1)-(o[b.status]!==undefined?o[b.status]:1);
    });
    return results;
  } catch(e) { Logger.log("getMyRequests Error: "+e); return []; }
}

function getAllRequests() {
  try {
    var sheet = getRequestsSheet();
    var data  = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    var results = [];
    for (var i = 1; i < data.length; i++) results.push(rowToRequest(data[i]));
    results.sort(function(a,b){
      var o={"รอดำเนินการ":0,"ไม่อนุมัติ":1,"อนุมัติ":2};
      return (o[a.status]!==undefined?o[a.status]:0)-(o[b.status]!==undefined?o[b.status]:0);
    });
    return results;
  } catch(e) { Logger.log("getAllRequests Error: "+e); return []; }
}

function approveRequest(reqId, approveBy, editedObj) {
  try {
    var sheet = getRequestsSheet();
    var data  = sheet.getDataRange().getValues();
    var now   = Utilities.formatDate(new Date(),"GMT+7","dd/MM/yyyy HH:mm:ss");
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString() !== reqId) continue;
      var requestBy = data[i][12] ? data[i][12].toString() : "";
      var staffList = Array.isArray(editedObj.staff)
        ? editedObj.staff
        : (editedObj.staff||"").split(",").map(function(s){ return s.trim(); }).filter(Boolean);
      var startStr = toDateStr(editedObj.start);
      var endStr   = toDateStr(editedObj.end);
      var sTimeStr = toTimeStr(editedObj.sTime) || "09:00";
      var eTimeStr = toTimeStr(editedObj.eTime) || "17:00";

      // ✅ fix overnight
      var fixed = fixOvernightJob(startStr, endStr, sTimeStr, eTimeStr);
      startStr = fixed.startStr;
      endStr   = fixed.endStr;

      // Conflict check
      if (staffList.length > 0) {
        var conflicts = checkStaffConflict(startStr, endStr, sTimeStr, eTimeStr, staffList, null);
        if (conflicts.length > 0) {
          var cm = conflicts.map(function(c){ return c.name+" (Job: "+c.jobId+" วันที่ "+c.date+")"; }).join(", ");
          return { success:false, msg:"⚠️ พบงานชน: "+cm };
        }
      }

      // อัปเดต Requests sheet
      sheet.getRange(i+1,6).setValue(startStr);
      sheet.getRange(i+1,7).setValue(endStr);
      sheet.getRange(i+1,8).setValue(sTimeStr);
      sheet.getRange(i+1,9).setValue(eTimeStr);
      sheet.getRange(i+1,10).setValue(editedObj.robots||"");
      sheet.getRange(i+1,11).setValue(staffList.join(", "));
      sheet.getRange(i+1,12).setValue(editedObj.remark||"");
      sheet.getRange(i+1,14).setValue("อนุมัติ");
      sheet.getRange(i+1,15).setValue(approveBy);
      sheet.getRange(i+1,16).setValue("");
      sheet.getRange(i+1,17).setValue(now);
      sheet.getRange(i+1,18).setValue(editedObj.jobType||"");

      // สร้าง Job
      var jobId     = "J-"+new Date().getTime();
      var dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
      var location  = (editedObj.location||"")+(editedObj.branch?" ("+editedObj.branch+")":"");
      var eventId   = "";

      // Calendar
      try {
        var cal = CalendarApp.getCalendarById(CALENDAR_ID);
        var ev  = cal.createEvent(
          "["+editedObj.cust+(editedObj.branch?" - "+editedObj.branch:"")+"] "+(editedObj.robots||""),
          fixed.sd, fixed.ed,
          { description:
              "📍 "+location+
              "\n👥 Staff: "+staffList.join(", ")+
              "\n📋 "+(editedObj.jobType||"-")+
              "\n📝 "+(editedObj.remark||"-")+
              "\n🔖 Job ID: "+jobId,
            location: location }
        );
        eventId = ev.getId();
        Logger.log("approveRequest Calendar created: "+eventId);
      } catch(e) { Logger.log("Calendar Error: "+e.toString()); }

      // บันทึก Data sheet
      dataSheet.appendRow([
        jobId, startStr, endStr, sTimeStr, eTimeStr,
        editedObj.cust, editedObj.branch||"", editedObj.jobType||"Service",
        editedObj.robots||"", "", staffList.join(", "), editedObj.remark||"",
        eventId, "งานใหม่", "", ""
      ]);

      // Email
      try {
        var emailMap = buildEmailMap();
        var reqEmail = emailMap[requestBy];
        if (reqEmail) MailApp.sendEmail(reqEmail, "✅ คำขอลงแผนงานได้รับการอนุมัติ",
          "สวัสดีคุณ "+requestBy+",\n\nคำขอของคุณได้รับการอนุมัติแล้ว\n\n━━━━━━━━━━━━━━━━━━━━\n"+
          "🔖 Req ID     : "+reqId+"\n🏢 ลูกค้า     : "+editedObj.cust+(editedObj.branch?" ("+editedObj.branch+")":"")+"\n"+
          "📋 ประเภทงาน  : "+(editedObj.jobType||"-")+"\n📅 วันที่      : "+startStr+" ถึง "+endStr+"\n"+
          "⏰ เวลา       : "+sTimeStr+" - "+eTimeStr+"\n🤖 หุ่นยนต์   : "+(editedObj.robots||"-")+"\n"+
          "👥 Staff       : "+staffList.join(", ")+"\n👤 อนุมัติโดย  : "+approveBy+"\n"+
          "━━━━━━━━━━━━━━━━━━━━\n\nRobot Operation System"
        );
        staffList.forEach(function(name) {
          var email = emailMap[name.trim()]; if(!email) return;
          MailApp.sendEmail(email, "📋 งานใหม่: "+editedObj.cust,
            "สวัสดีคุณ "+name+",\n\nคุณได้รับมอบหมายงานใหม่\n\n━━━━━━━━━━━━━━━━━━━━\n"+
            "🔖 Job ID     : "+jobId+"\n🏢 ลูกค้า     : "+editedObj.cust+(editedObj.branch?" ("+editedObj.branch+")":"")+"\n"+
            "📋 ประเภทงาน  : "+(editedObj.jobType||"-")+"\n📍 สถานที่    : "+location+"\n"+
            "📅 วันที่      : "+startStr+" ถึง "+endStr+"\n⏰ เวลา       : "+sTimeStr+" - "+eTimeStr+"\n"+
            "🤖 หุ่นยนต์   : "+(editedObj.robots||"-")+"\n"+
            "━━━━━━━━━━━━━━━━━━━━\n\nRobot Operation System"
          );
        });
      } catch(e) { Logger.log("Email Error: "+e.toString()); }

      return { success:true, msg:"อนุมัติและลงแผนงานสำเร็จ! ✅" };
    }
    return { success:false, msg:"ไม่พบคำขอ reqId: "+reqId };
  } catch(e) {
    Logger.log("approveRequest Error: "+e.toString());
    return { success:false, msg:"approveRequest Error: "+e.toString() };
  }
}

function rejectRequest(reqId, approveBy, reason) {
  try {
    var sheet = getRequestsSheet();
    var data  = sheet.getDataRange().getValues();
    var now   = Utilities.formatDate(new Date(),"GMT+7","dd/MM/yyyy HH:mm:ss");
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString() !== reqId) continue;
      var requestBy = data[i][12] ? data[i][12].toString() : "";
      sheet.getRange(i+1,14).setValue("ไม่อนุมัติ");
      sheet.getRange(i+1,15).setValue(approveBy);
      sheet.getRange(i+1,16).setValue(reason);
      sheet.getRange(i+1,17).setValue(now);
      try {
        var emailMap = buildEmailMap();
        var reqEmail = emailMap[requestBy];
        if (reqEmail) MailApp.sendEmail(reqEmail,"❌ คำขอลงแผนงานไม่ได้รับการอนุมัติ",
          "สวัสดีคุณ "+requestBy+",\n\nคำขอของคุณไม่ได้รับการอนุมัติ\n\n━━━━━━━━━━━━━━━━━━━━\n"+
          "🔖 Req ID    : "+reqId+"\n❌ เหตุผล   : "+(reason||"-")+"\n"+
          "👤 ปฏิเสธโดย: "+approveBy+"\n━━━━━━━━━━━━━━━━━━━━\n\nคุณสามารถแก้ไขและส่งคำขอใหม่ได้\n\nRobot Operation System"
        );
      } catch(e) { Logger.log("Email Error: "+e); }
      return { success:true, msg:"ปฏิเสธคำขอเรียบร้อย" };
    }
    return { success:false, msg:"ไม่พบคำขอ" };
  } catch(e) { return { success:false, msg:e.toString() }; }
}

function resubmitRequest(reqId, obj) {
  try {
    var sheet = getRequestsSheet();
    var data  = sheet.getDataRange().getValues();
    var now   = Utilities.formatDate(new Date(),"GMT+7","dd/MM/yyyy HH:mm:ss");
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString() !== reqId) continue;
      sheet.getRange(i+1,3).setValue(obj.cust);
      sheet.getRange(i+1,4).setValue(obj.branch);
      sheet.getRange(i+1,5).setValue(obj.location);
      sheet.getRange(i+1,6).setValue(obj.start);
      sheet.getRange(i+1,7).setValue(obj.end);
      sheet.getRange(i+1,8).setValue("'" + (obj.sTime||"09:00"));
      sheet.getRange(i+1,9).setValue("'" + (obj.eTime||"17:00"));
      sheet.getRange(i+1,10).setValue(obj.robots);
      sheet.getRange(i+1,12).setValue(obj.remark||"");
      sheet.getRange(i+1,14).setValue("รอดำเนินการ");
      sheet.getRange(i+1,15).setValue("");
      sheet.getRange(i+1,16).setValue("");
      sheet.getRange(i+1,17).setValue(now);
      sheet.getRange(i+1,18).setValue(obj.jobType||"");
      return { success:true, msg:"ส่งคำขอใหม่สำเร็จ! ✅" };
    }
    return { success:false, msg:"ไม่พบคำขอ" };
  } catch(e) { return { success:false, msg:e.toString() }; }
}

// ============================================================
// USERS
// ============================================================
function createUser(obj) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    var data  = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++)
      if (data[i][0].toString().trim().toLowerCase() === obj.username.trim().toLowerCase())
        return { success:false, msg:"❌ Username \""+obj.username+"\" มีอยู่แล้ว" };
    sheet.appendRow([obj.username, obj.password, obj.fullname, obj.email, obj.role, obj.dept]);
    return { success:true, msg:"สร้าง User สำเร็จ!" };
  } catch(e) { return { success:false, msg:e.toString() }; }
}

function updateUser(obj) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    var data  = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === obj.username.trim()) {
        if (obj.password) sheet.getRange(i+1,2).setValue(obj.password);
        sheet.getRange(i+1,3).setValue(obj.fullname);
        sheet.getRange(i+1,4).setValue(obj.email);
        sheet.getRange(i+1,5).setValue(obj.role);
        sheet.getRange(i+1,6).setValue(obj.dept);
        return { success:true, msg:"อัปเดต User สำเร็จ!" };
      }
    }
    return { success:false, msg:"ไม่พบ Username นี้" };
  } catch(e) { return { success:false, msg:e.toString() }; }
}

function deleteUser(username) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    var data  = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === username.trim()) {
        sheet.deleteRow(i+1);
        return { success:true, msg:"ลบ User สำเร็จ!" };
      }
    }
    return { success:false, msg:"ไม่พบ Username นี้" };
  } catch(e) { return { success:false, msg:e.toString() }; }
}

function getUserList() {
  try {
    var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users').getDataRange().getValues();
    if (data.length <= 1) return [];
    return data.slice(1).map(function(row) {
      return { username:row[0]||"", fullname:row[2]||"", email:row[3]||"", role:row[4]||"", dept:row[5]||"" };
    });
  } catch(e) { return []; }
}

function testCheckEventId() {
  var data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Data').getDataRange().getValues();
  data.slice(1).forEach(function(row) {
    Logger.log("JobID: "+row[0]+" | EventID: ["+row[12]+"] | Start: "+row[1]);
  });
}

// ============================================================
// DASHBOARD
// ============================================================
function getDashboardData(filter) {
  try {
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var data = ss.getSheetByName('Data').getDataRange().getValues();
    if (data.length <= 1) return emptyDashboard();

    // โหลด JobTypes จาก Sheet
    var jobTypes = [];
    try {
      var jtSheet = ss.getSheetByName('JobTypes');
      if (jtSheet) jobTypes = jtSheet.getDataRange().getValues().slice(1)
        .map(function(r){ return r[0]?r[0].toString().trim():""; }).filter(Boolean);
    } catch(e){}

    // โหลด staff→dept map
    var staffDeptMap = {};
    ss.getSheetByName('Users').getDataRange().getValues().slice(1).forEach(function(u) {
      var name = (u[2]||"").toString().trim();
      var dept = (u[5]||"Other").toString().trim();
      if (name) staffDeptMap[name] = dept;
    });

    var rows = data.slice(1).map(function(row) {
      return {
        id:      (row[0]||"").toString(),
        start:   row[1] ? toDateStr(row[1]) : "",
        end:     row[2] ? toDateStr(row[2]) : "",
        cust:    (row[5]||"").toString(),
        branch:  (row[6]||"").toString(),
        jobType: (row[7]||"").toString(),
        staff:   (row[10]||"").toString(),
        status:  (row[13]||"งานใหม่").toString()
      };
    }).filter(function(r){ return r.start !== ""; });

    // ── Apply filters ──
    var f = filter || {};
    if (f.dateFrom) rows = rows.filter(function(r){ return r.start >= f.dateFrom; });
    if (f.dateTo)   rows = rows.filter(function(r){ return r.start <= f.dateTo; });
    if (f.jobType)  rows = rows.filter(function(r){ return r.jobType === f.jobType; });
    if (f.cust)     rows = rows.filter(function(r){ return r.cust === f.cust; });
    if (f.staff)    rows = rows.filter(function(r){
      return r.staff.split(",").map(function(s){ return s.trim(); }).indexOf(f.staff) > -1;
    });
    if (f.dept) {
      var deptStaff = Object.keys(staffDeptMap).filter(function(n){ return staffDeptMap[n] === f.dept; });
      rows = rows.filter(function(r){
        var arr = r.staff.split(",").map(function(s){ return s.trim(); });
        return arr.some(function(s){ return deptStaff.indexOf(s) > -1; });
      });
    }

    // ── KPI ──
    var kpi = {
      total:      rows.length,
      done:       rows.filter(function(r){ return r.status === "เสร็จสิ้น"; }).length,
      inProgress: rows.filter(function(r){ return r.status === "กำลังดำเนินการ"; }).length,
      newJob:     rows.filter(function(r){ return r.status === "งานใหม่"; }).length
    };

    // ── Trend by day ──
    var trendDay = {}, trendWeek = {}, trendMonth = {}, trendYear = {};
    rows.forEach(function(r) {
      if (!r.start) return;
      var d   = new Date(r.start);
      var day = r.start;
      var wk  = (function(){
        var tmp = new Date(d);
        tmp.setDate(d.getDate() - (d.getDay()===0?6:d.getDay()-1));
        return tmp.toISOString().substring(0,10);
      })();
      var mo  = r.start.substring(0,7);
      var yr  = r.start.substring(0,4);
      trendDay[day]   = (trendDay[day]  ||0) + 1;
      trendWeek[wk]   = (trendWeek[wk]  ||0) + 1;
      trendMonth[mo]  = (trendMonth[mo] ||0) + 1;
      trendYear[yr]   = (trendYear[yr]  ||0) + 1;
    });

    function toTrend(map) {
      var keys = Object.keys(map).sort();
      return { labels: keys, values: keys.map(function(k){ return map[k]; }) };
    }

    // ── Staff stats ──
    var staffMap = {};
    rows.forEach(function(r) {
      if (!r.staff) return;
      r.staff.split(",").forEach(function(s) {
        s = s.trim(); if (!s) return;
        if (!staffMap[s]) staffMap[s] = { name:s, dept:staffDeptMap[s]||"Other", total:0, done:0, inProgress:0, newJob:0 };
        staffMap[s].total++;
        if (r.status === "เสร็จสิ้น")           staffMap[s].done++;
        else if (r.status === "กำลังดำเนินการ") staffMap[s].inProgress++;
        else                                     staffMap[s].newJob++;
      });
    });
    var staffStats = Object.values(staffMap).sort(function(a,b){ return b.total - a.total; });

    // ── Job type stats ──
    var typeMap = {};
    rows.forEach(function(r) {
      var t = r.jobType || "ไม่ระบุ";
      typeMap[t] = (typeMap[t]||0) + 1;
    });
    var jobTypeStats = Object.keys(typeMap).map(function(k){ return { label:k, value:typeMap[k] }; })
      .sort(function(a,b){ return b.value - a.value; });

    // ── Customer stats ──
    var custMap = {};
    rows.forEach(function(r) {
      var c = r.cust || "ไม่ระบุ";
      if (!custMap[c]) custMap[c] = { total:0, byType:{} };
      custMap[c].total++;
      var t = r.jobType||"ไม่ระบุ";
      custMap[c].byType[t] = (custMap[c].byType[t]||0) + 1;
    });
    var custStats = Object.keys(custMap).map(function(k){
      return { name:k, total:custMap[k].total, byType:custMap[k].byType };
    }).sort(function(a,b){ return b.total - a.total; });

    // ── Heatmap ──
    var dowLabels = ["จ","อ","พ","พฤ","ศ","ส","อา"];
    var heatMap = {};
    rows.forEach(function(r) {
      if (!r.start || !r.staff) return;
      var dow = new Date(r.start).getDay();
      dow = dow === 0 ? 6 : dow - 1;
      r.staff.split(",").forEach(function(s) {
        s = s.trim(); if (!s) return;
        if (!heatMap[s]) heatMap[s] = [0,0,0,0,0,0,0];
        heatMap[s][dow]++;
      });
    });

    // ── Dept summary ──
    var deptMap = {};
    rows.forEach(function(r) {
      if (!r.staff) return;
      r.staff.split(",").forEach(function(s) {
        s = s.trim(); if (!s) return;
        var dept = staffDeptMap[s]||"Other";
        deptMap[dept] = (deptMap[dept]||0) + 1;
      });
    });
    var deptStats = Object.keys(deptMap).map(function(k){ return { dept:k, total:deptMap[k] }; })
      .sort(function(a,b){ return b.total - a.total; });

    // ── Dept list & Staff list & Cust list — จาก rows ที่ filter วันที่แล้วเท่านั้น ──
    var activeStaff = {};
    rows.forEach(function(r) {
      if (!r.staff) return;
      r.staff.split(",").forEach(function(s) {
        s = s.trim(); if (!s) return;
        activeStaff[s] = staffDeptMap[s] || "Other";
      });
    });

    var activeDepts = [];
    Object.values(activeStaff).forEach(function(d) {
      if (activeDepts.indexOf(d) < 0) activeDepts.push(d);
    });
    activeDepts.sort();

    var filteredStaffList = Object.keys(activeStaff).filter(function(n) {
      if (f.dept) return activeStaff[n] === f.dept;
      return true;
    }).sort();

    var activeCusts = [];
    rows.forEach(function(r) {
      if (r.cust && activeCusts.indexOf(r.cust) < 0) activeCusts.push(r.cust);
    });
    activeCusts.sort();

    var activeJobTypes = [];
    rows.forEach(function(r) {
      if (r.jobType && activeJobTypes.indexOf(r.jobType) < 0) activeJobTypes.push(r.jobType);
    });
    if (jobTypes.length > 0) {
      activeJobTypes.sort(function(a,b) {
        return jobTypes.indexOf(a) - jobTypes.indexOf(b);
      });
    } else {
      activeJobTypes.sort();
    }

    return {
      kpi, staffStats,
      trendDay:   toTrend(trendDay),
      trendWeek:  toTrend(trendWeek),
      trendMonth: toTrend(trendMonth),
      trendYear:  toTrend(trendYear),
      jobTypeStats, custStats,
      heatMap, dowLabels, deptStats,
      deptList:  activeDepts,
      staffList: filteredStaffList,
      custList:  activeCusts,
      jobTypes:  activeJobTypes
    };
  } catch(e) {
    Logger.log("getDashboardData Error: "+e);
    return emptyDashboard();
  }
}
function emptyDashboard() {
  return {
    kpi: { total:0, done:0, inProgress:0, newJob:0 },
    staffStats: [], trend: { labels:[], values:[] },
    jobTypeStats: [], custStats: [], heatMap: {}, dowLabels: [], deptStats: []
  };
}

function testConflictCheck() {
  // ใส่ข้อมูลจริงที่กำลัง approve
  var start     = "2026-03-16";
  var end       = "2026-03-16";
  var sTime     = "09:00";
  var eTime     = "17:00";
  var staffList = ["ใส่ชื่อ Staff ที่ขึ้นว่าชน"]; // Full Name จริง

  var conflicts = checkStaffConflict(start, end, sTime, eTime, staffList, null);
  Logger.log("Conflicts found: " + conflicts.length);
  conflicts.forEach(function(c) {
    Logger.log("Name: "+c.name+" | JobID: "+c.jobId+" | Date: "+c.date);
  });

  // ดูงานทั้งหมดวันที่ 16/3
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data').getDataRange().getValues();
  data.slice(1).forEach(function(row) {
    var d = row[1] ? new Date(row[1]) : null;
    if (!d) return;
    if (d.getMonth()===2 && d.getDate()===16) { // March = 2
      Logger.log("Job: "+row[0]+" Staff: ["+row[10]+"] Start: "+row[1]+" End: "+row[2]);
    }
  });
}

function testUserMap() {
  var data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Users').getDataRange().getValues();
  Logger.log("Headers: " + JSON.stringify(data[0]));
  data.slice(1).forEach(function(u) {
    Logger.log("col A(username): ["+u[0]+"] | col C(fullname): ["+u[2]+"] | col F(dept): ["+u[5]+"]");
  });
  
  // ดู Data sheet staff column
  var jobs = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Data').getDataRange().getValues();
  Logger.log("--- Data Sheet Staff ---");
  jobs.slice(1).forEach(function(row) {
    Logger.log("JobID: "+row[0]+" | Staff col K: ["+row[10]+"]");
  });
}

function testBusyOnDate() {
  // วันที่ที่กำลัง approve
  var start  = "2026-03-16";
  var end    = "2026-03-23";
  var sTime  = "10:17";
  var eTime  = "17:17";

  var result = getAvailableStaffByDept(start, end, sTime, eTime);
  Object.keys(result).forEach(function(dept) {
    Logger.log("=== " + dept + " ===");
    result[dept].forEach(function(s) {
      Logger.log(s.name + " → " + (s.available ? "ว่าง" : "ติดงาน"));
    });
  });
}

function testWhyBusy() {
  var checkName = "thanapong"; // เปลี่ยนชื่อได้
  var reqStart  = new Date("2026-03-16T10:17:00");
  var reqEnd    = new Date("2026-03-23T17:17:00");

  var data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Data').getDataRange().getValues();

  data.slice(1).forEach(function(row) {
    if (!row[10]) return;
    var staff = row[10].toString().split(",").map(function(s){ return s.trim().toLowerCase(); });
    if (staff.indexOf(checkName.toLowerCase()) === -1) return;

    var rs = toDateStr(row[1]);
    var re = toDateStr(row[2]);
    var st = toTimeStr(row[3]) || "00:00";
    var et = toTimeStr(row[4]) || "23:59";
    var rowStart = new Date(rs+"T"+st+":00");
    var rowEnd   = new Date(re+"T"+et+":00");

    var overlap = !(reqStart >= rowEnd || reqEnd <= rowStart);
    Logger.log(
      "JobID: "+row[0]+
      " | "+rs+" "+st+" → "+re+" "+et+
      " | overlap: "+overlap
    );
  });
}

function testFindGhostJobs() {
  var cal  = CalendarApp.getCalendarById("3vsvb0a1nblh7odvi0qecsb0a0@group.calendar.google.com");
  var data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Data').getDataRange().getValues();

  data.slice(1).forEach(function(row) {
    var jobId   = row[0].toString();
    var eventId = row[12] ? row[12].toString().trim() : "";
    var staff   = row[10] ? row[10].toString() : "";
    var start   = row[1]  ? toDateStr(row[1])  : "";

    if (!eventId) {
      Logger.log("NO EventID | JobID: "+jobId+" | Start: "+start+" | Staff: "+staff);
      return;
    }
    try {
      var ev = cal.getEventById(eventId);
      if (!ev) Logger.log("EVENT DELETED | JobID: "+jobId+" | Start: "+start+" | Staff: "+staff);
    } catch(e) {
      Logger.log("EVENT ERROR | JobID: "+jobId+" | "+e.toString());
    }
  });
  Logger.log("Done");
}

function testJobDetail() {
  var data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Data').getDataRange().getValues();
  
  data.slice(1).forEach(function(row) {
    if (row[0].toString() === "J-1772703420874") {
      Logger.log("JobID: "   + row[0]);
      Logger.log("Customer: "+ row[5] + " " + row[6]);
      Logger.log("Start: "   + toDateStr(row[1]) + " " + toTimeStr(row[3]));
      Logger.log("End: "     + toDateStr(row[2]) + " " + toTimeStr(row[4]));
      Logger.log("Staff: "   + row[10]);
      Logger.log("Status: "  + row[13]);
    }
  });
}

function testCheckRequest() {
  var data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Requests').getDataRange().getValues();
  var last = data[data.length-1];
  Logger.log("ReqID: "  + last[0]);
  Logger.log("Start: "  + last[5]);
  Logger.log("End: "    + last[6]);
  Logger.log("sTime: "  + last[7]);
  Logger.log("eTime: "  + last[8]);
  Logger.log("Status: " + last[13]);
}
function testPendingRequests() {
  var data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Requests').getDataRange().getValues();
  
  data.slice(1).forEach(function(row) {
    if (row[13].toString() !== "รอดำเนินการ") return;
    Logger.log("ReqID: " + row[0]);
    Logger.log("Start: " + row[5]);
    Logger.log("End: "   + row[6]);
    Logger.log("sTime: " + row[7]);
    Logger.log("eTime: " + row[8]);
    Logger.log("---");
  });
}

function getJobTypes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('JobTypes');
  if (!sh) return [];
  return sh.getRange(2, 1, Math.max(sh.getLastRow()-1, 1), 1)
           .getValues()
           .map(r => r[0].toString().trim())
           .filter(Boolean);
}

function getRequestsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Requests');
  if (!sheet) {
    sheet = ss.insertSheet('Requests');
    sheet.appendRow(['Req_ID','Date','Customer','Branch','Location','Start','End',
      'sTime','eTime','Robots','Staff','Remark','RequestBy','Status',
      'ApproveBy','Reason','UpdateTime','JobType']);
    // ✅ บังคับ column sTime(H) และ eTime(I) เป็น Plain Text
    sheet.getRange('H:H').setNumberFormat('@');
    sheet.getRange('I:I').setNumberFormat('@');
  }
  return sheet;
}