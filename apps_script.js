// Baldwin MLS Perchwell CRM — Google Apps Script Backend
// Deploy as Web App: Execute as "Me", Access "Anyone"

const SHEET_NAME = "Members";
const LOG_SHEET = "Interactions";

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const params = e.parameter;
  const action = params.action;
  
  try {
    let result;
    if (action === "getMembers") result = getMembers();
    else if (action === "updateMember") result = updateMember(params);
    else if (action === "addInteraction") result = addInteraction(params);
    else if (action === "getInteractions") result = getInteractions(params.memberId);
    else result = { error: "Unknown action" };
    
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getMembers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = initSheet(ss);
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const members = data.slice(1).map((row, i) => {
    const obj = {};
    headers.forEach((h, j) => obj[h] = row[j]);
    obj._row = i + 2;
    return obj;
  }).filter(m => m.id !== "");
  
  return { members };
}

function updateMember(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const idCol = headers.indexOf("id");
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(params.id)) {
      const fields = ["status", "trainingDate", "trainingType", "notes", "issues"];
      fields.forEach(f => {
        const col = headers.indexOf(f);
        if (col >= 0 && params[f] !== undefined) {
          sheet.getRange(i + 1, col + 1).setValue(params[f]);
        }
      });
      return { success: true };
    }
  }
  return { error: "Member not found" };
}

function addInteraction(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName(LOG_SHEET);
  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_SHEET);
    logSheet.getRange(1,1,1,6).setValues([["memberId","memberName","type","date","loggedBy","note"]]);
    logSheet.getRange(1,1,1,6).setFontWeight("bold");
    logSheet.setFrozenRows(1);
  }
  logSheet.appendRow([params.memberId, params.memberName, params.type, params.date, params.loggedBy, params.note]);
  return { success: true };
}

function getInteractions(memberId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(LOG_SHEET);
  if (!logSheet) return { interactions: [] };
  
  const data = logSheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf("memberId");
  
  const interactions = data.slice(1)
    .filter(row => String(row[idCol]) === String(memberId))
    .map(row => {
      const obj = {};
      headers.forEach((h, j) => obj[h] = row[j]);
      return obj;
    });
  
  return { interactions };
}

function initSheet(ss) {
  const sheet = ss.insertSheet(SHEET_NAME);
  const headers = ["id","name","role","office","email","phone","status","trainingDate","trainingType","notes","issues"];
  sheet.getRange(1,1,1,headers.length).setValues([headers]);
  sheet.getRange(1,1,1,headers.length).setFontWeight("bold");
  sheet.setFrozenRows(1);
  
  // Pre-load seed data
  const seed = getSeedData();
  if (seed.length > 0) {
    const rows = seed.map((m,i) => [
      i, m.name, m.role, m.office, m.email, m.phone,
      (m.notes || m.issues) ? "In outreach" : "Not contacted",
      "", "", m.notes, m.issues
    ]);
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);
  return sheet;
}

function getSeedData() {
  return [
    {name:"Miranda Baldridge",role:"MLS BOD",office:"ADVANTAGE REAL ESTATE ROBERTSDALE",email:"miranda@gulffrontproperties.com",phone:"251-284-1068",notes:"",issues:""},
    {name:"Bowen Weir",role:"Association BOD",office:"B4 PROPERTIES, LLC",email:"bowen@b4prop.com",phone:"251.270.9099",notes:"Monday 3/9 to review the CMA workflow. Bowen uses the CMA Quick View in Paragon, so we showed the Quick CMA flow in Perchwell. 3/23: Asked about the export from search.",issues:"Had a hard time logging in and remembering password for BRs portal. 95% commercial. Statistical reporting of who is selling what."},
    {name:"Terryl Reeves",role:"MLS BOD, President-Elect",office:"BELLATOR REAL ESTATE, LLC",email:"treeves@gobellator.com",phone:"251-29-7484",notes:"Looped into meetings with Product & Diane + Francine prior to kick-off. Has great energy. Very excited, expressed interest in providing additional feedback.",issues:""},
    {name:"Kristen Meador",role:"Association BOD",office:"BELLATOR REAL ESTATE, LLC",email:"kbmeador@gobellator.com",phone:"251 459-1546",notes:"Sat with her at lunch.",issues:""},
    {name:"Sondra Blackwell",role:"Association BOD",office:"BLACKWELL REALTY INC",email:"sondra@blackwellrealtyinc.com",phone:"251-455-1172",notes:"Often co-lists with her husband. Sat with her at lunch.",issues:""},
    {name:"Constance Calambakas",role:"Association BOD",office:"BRETT ROBINSON",email:"constancec@brettrobinson.com",phone:"251-407-6187",notes:"",issues:""},
    {name:"Brooke Butler",role:"Association BOD",office:"BUTLER & CO REAL ESTATE",email:"brooke@thebutlerandco.com",phone:"251-232-9901",notes:"",issues:""},
    {name:"Francine Carstensen",role:"MLS BOD, Current President",office:"ELITE REAL ESTATE SOLUTIONS, LLC",email:"francinesellsbaldwin@gmail.com",phone:"251-213-3522",notes:"",issues:""},
    {name:"Diane Martino",role:"MLS BOD, Immediate Past President",office:"EXIT REALTY LANDMARK",email:"dnj.exitrealtylandmark@gmail.com",phone:"251-259-8037",notes:"",issues:""},
    {name:"Renee Marshall",role:"Association BOD",office:"EXIT REALTY LYON",email:"renee.exitrealtylyon@gmail.com",phone:"251-421-2614",notes:"",issues:""},
    {name:"Milan Portis",role:"MLS BOD",office:"KELLER WILLIAMS ALABAMA GULF COAST",email:"milanportis@kw.com",phone:"251-422-2761",notes:"Had a lot of questions during the on-site workflow walkthrough.",issues:""},
    {name:"Stephanie Jenkins",role:"Association BOD",office:"LIVING MY BEST LIFE REALTY",email:"stephanie@livingmybestliferealty.com",phone:"305-916-0812",notes:"Joined 3/5 office hours. Fan of what she has seen, interested in being an early adopter for new functionality.",issues:""},
    {name:"Susan Shallow",role:"Association BOD & MLS BOD",office:"RE/MAX PARADISE",email:"susan@susanshallow.com",phone:"251-797-0501",notes:"",issues:""},
    {name:"Fallon Young",role:"MLS BOD",office:"RE/MAX PARADISE",email:"f.young@alabamaparadise.com",phone:"251-259-8132",notes:"Joined 3/6 office hours, focused on search, strong reaction to property type changes. Met 3/10 re: CMA flow. 3/13: Discussed feedback after 2 weeks. Loves the dashboard, wants faster search filtering.",issues:"Wants to go see gators next time. Does short-term rental homes. Would help on mortgage/financing tool recommendations."},
    {name:"Walter Billups",role:"Association BOD",office:"THE COLONY AT THE GRAND REALTY",email:"wbillups@ingramnewhomes.com",phone:"228-216-7889",notes:"Very excited about Perchwell and particularly the mobile app.",issues:""},
    {name:"Beverley Valrie",role:"MLS BOD",office:"VALPOINTE REAL ESTATE & DEVELOPMENT",email:"beverleyvalrie@gmail.com",phone:"251-644-8201",notes:"Joined 3/6 office hours. Asked about CMA, showed Quick CMA & setup. NOTE: Name sometimes misspelled as Beverly — pay attention in comms.",issues:""},
    {name:"Kendall Wahlert",role:"Association BOD",office:"WATERS EDGE REALTY",email:"kendallsoldit@gmail.com",phone:"706-575-8800",notes:"Not in person for BOD kick-off, had a 1-1 over Zoom later. Raised the J. Larry Newton elementary school issue.",issues:""},
    {name:"Rich Caldwell",role:"PAB",office:"Coldwell Banker Coastal Realty-Foley",email:"SRichCaldwell@gmail.com",phone:"251-504-6791",notes:"Adrian & Rich working on reports & analytics. Reviewing Month End Report from Paragon on Fri Mar 13.",issues:""},
    {name:"David Kahalley",role:"PAB",office:"Shamrock Properties, LLC",email:"david@shamrockhouses.com",phone:"251-533-9033",notes:"Meeting on Mar 12 to discuss reports/analytics, client collaboration, and feedback on mobile.",issues:""},
    {name:"Andrea Shilston",role:"PAB",office:"Kaiser Sotheby's Int",email:"andrea@kaisersir.com",phone:"251-752-0192",notes:"",issues:""},
    {name:"Alison Ward",role:"PAB",office:"Mobile Bay Realty",email:"alison@mobilebayrealty.com",phone:"850-485-3600",notes:"Discussed CMAs on Mar 12. Entire office does not use Paragon CMA. Strong feedback around starting with a branded PDF including price range chart.",issues:""},
    {name:"Ginny Stopa",role:"PAB",office:"NextHome Gulf Coast Living",email:"ginny@ginnystopa.com",phone:"251-623-0200",notes:"Demo'd Perchwell at sales meeting Wed Mar 4. Follow-up questions on Teams, commercial listings, attached/detached searching.",issues:""},
    {name:"Amy Slade",role:"PAB",office:"NextHome Gulf Coast Living",email:"amy.h.slade@gmail.com",phone:"850-420-3413",notes:"",issues:""},
    {name:"Holly Hayek",role:"PAB",office:"Phoenix Properties",email:"hollyhuntshomes@gmail.com",phone:"251-550-1121",notes:"",issues:""},
    {name:"Tommy Stanton",role:"PAB",office:"Exit Orange Beach Realty",email:"Tommy.exit@gmail.com",phone:"251-213-1860",notes:"",issues:""},
    {name:"Jennifer Foutch",role:"Association President",office:"",email:"",phone:"",notes:"",issues:""},
    {name:"Wendy Alley",role:"MLS Staff",office:"",email:"Wendy@baldwinrealtors.com",phone:"(251) 270-2723",notes:"20 years of teaching experience.",issues:"Pain point: SSMs not knowing which features different MLS have."},
    {name:"Garet Ikner",role:"MLS Staff",office:"",email:"garet@baldwinrealtors.com",phone:"D: (251) 270-2777",notes:"Started as IT, moved into this industry 6 years ago. Everything escalated to him (bottleneck). Handles a lot of the feeds.",issues:""},
    {name:"Dusti Hafner",role:"MLS Staff",office:"",email:"Dusti@baldwinrealtors.com",phone:"D: 251-270-2725",notes:"Teaches member classes currently but struggling to pick up on Perchwell.",issues:""},
    {name:"Meagan Smith",role:"MLS Staff",office:"",email:"meagan@baldwinrealtors.com",phone:"D: (251) 270-2779",notes:"Has not taught classes before, needs more work.",issues:""},
    {name:"David Cowles",role:"MLS Staff",office:"",email:"David@baldwinrealtors.com",phone:"D: (251) 616-9769",notes:"50% MLS / 50% Association. Never taught members before. Strongest trainer so far.",issues:""},
    {name:"Allison",role:"MLS Staff",office:"",email:"",phone:"",notes:"LMS meeting attended.",issues:""},
    {name:"Brenda Copeland",role:"Broker, UAT",office:"Ashurst & Niemeyer LLC",email:"brenda@ashurstandniemeyer.com",phone:"251-709-9033",notes:"",issues:"Follow up when Android is ready. Really liked the one pager printed. Follow up with PDF version."},
    {name:"Craig Jackson",role:"Member",office:"",email:"cjackson@mindspring.com",phone:"251-923-6436",notes:"Spent a lot of time with him. Wendy mentioned he calls quite a bit. Expressed frustration about Paragon removing the Area filter. Worked with him on a workaround, seemed happy.",issues:""},
    {name:"Judd Gillespie",role:"Member",office:"",email:"judd@askjudd.com",phone:"251-895-3434",notes:"Office hours attended. Ran through activities and was very engaged.",issues:""},
    {name:"Anne Bodet",role:"Member",office:"",email:"exitanneb@gmail.com",phone:"251-599-8175",notes:"",issues:"Was missing add/edit and was very impressed when we fixed it in session."},
    {name:"Melissa Musick",role:"Member",office:"",email:"melissamusick@kw.com",phone:"251-269-0545",notes:"",issues:""},
    {name:"Robbie Irvine",role:"Member",office:"",email:"rlirvine@llbb.com",phone:"251-510-5818",notes:"",issues:""},
    {name:"Miranda Zavison",role:"Member",office:"",email:"mirandaz@pointesouth.com",phone:"850-417-3739",notes:"",issues:""},
    {name:"Kristy Bushaw",role:"Member",office:"",email:"kristy@kristybushaw.com",phone:"251-223-1111",notes:"",issues:""},
    {name:"Catherine Quinley",role:"Member",office:"",email:"catherinequinley@robertsbrothers.com",phone:"251-709-1040",notes:"Some of her brokerage is also Mobile. Had permission issues. Known in her office for finding bugs.",issues:""},
    {name:"Samantha Yeager",role:"Member",office:"",email:"samanthayeager@kw.com",phone:"251-550-1234",notes:"",issues:""},
    {name:"Greg Maples",role:"Member",office:"",email:"greg@themaplescompany.com",phone:"251-634-0714",notes:"",issues:""},
    {name:"Leon Nelson Jr",role:"Member",office:"",email:"leon@nelsonappraisals.net",phone:"251-232-2688",notes:"",issues:""},
    {name:"Jan Volovecky",role:"Member",office:"",email:"janwithcare@gmail.com",phone:"251-518-6638",notes:"",issues:""},
    {name:"Jennifer Lawrence",role:"Member",office:"",email:"jenlawrealtor@gmail.com",phone:"251-233-8623",notes:"",issues:""},
    {name:"Winford McGaster",role:"Member",office:"",email:"soldbywinford@gmail.com",phone:"251-979-8832",notes:"",issues:""},
    {name:"Donna Gardner",role:"Member",office:"",email:"DonnaGisMyAgent@gmail.com",phone:"251-622-3688",notes:"Most of training was easy for her. Had a few questions about analytics.",issues:""},
    {name:"Courtney Cathers",role:"Member",office:"",email:"courtneycathers1@gmail.com",phone:"251-367-0760",notes:"",issues:""},
    {name:"Heather Cole",role:"Member",office:"",email:"heather@dustycoleteam.com",phone:"251-978-8600",notes:"",issues:""},
    {name:"Jake Devantier",role:"Member",office:"",email:"jake@localpropertyinc.com",phone:"901-485-8426",notes:"Office hours attended.",issues:""},
    {name:"Brianna Johnson",role:"Member",office:"",email:"bjohnson@gobellator.com",phone:"251-250-9516",notes:"",issues:""},
    {name:"Andrea Torbert",role:"Member",office:"",email:"Andrea@cbcoastalrealty.com",phone:"251-747-7196",notes:"Sat next to Craig. Focused on the hot sheet and wanted a bar showing open houses.",issues:""},
    {name:"Evelyn Hall",role:"Member",office:"",email:"evelynhallhomes@gmail.com",phone:"251-232-0858",notes:"Very excited for Perchwell. Loved the search and mobile.",issues:""},
    {name:"Laura Jones",role:"Member",office:"",email:"ljones@goBellator.com",phone:"251-928-0031",notes:"",issues:""},
    {name:"Sandy Davenport",role:"Member",office:"",email:"sandybeachrealtor@yahoo.com",phone:"251-550-5500",notes:"Sat next to Wendy and shared a bunch of feedback.",issues:""},
    {name:"Ben Duhon",role:"Member",office:"",email:"benduhon@kw.com",phone:"251-680-6898",notes:"",issues:""},
    {name:"Bret Jennings",role:"Member",office:"",email:"BretJennings@kw.com",phone:"256-366-0665",notes:"",issues:""},
    {name:"Angie Hawes",role:"Member",office:"",email:"angie@ashurstandniemeyer.com",phone:"251-517-5419",notes:"",issues:""},
    {name:"Kimberly Moody",role:"Member",office:"",email:"kimberlyomoody@aol.com",phone:"251-599-7178",notes:"",issues:""},
    {name:"Greg Foote",role:"Member",office:"",email:"greg@gregoryfooterealestate.com",phone:"251-504-9224",notes:"",issues:""},
    {name:"Patsy Miller",role:"Member",office:"",email:"PatsyMiller@BlueHeron-Realty.com",phone:"(251) 402-7170",notes:"Office hours attended. Very nervous about the conversion, will want to attend as many training sessions as possible. Needs hand-holding.",issues:""},
    {name:"Sadie Bonilla",role:"Member",office:"",email:"s.bonilla@alabamaparadise.com",phone:"",notes:"",issues:""}
  ];
}
