const ss = SpreadsheetApp.getActiveSpreadsheet();
let sheet = ss.getSheetByName("Member");
function main() {
  const members = getMember();
  members.forEach((member) => {
    createAndActiveSheet(member);
    let MemberDataSet = getChData(member);
    insertRow(MemberDataSet);
  });
}
function getMember() {
  sheet = ss.getSheetByName("Member");
  const lastRow = sheet.getLastRow();
  let memberList = [];
  for (let i = 2; i <= lastRow; i++) {
    if (sheet.getRange(i, 1).getValue()) {
      memberList.push({
        name: sheet.getRange(i, 1).getValue(),
        id: sheet.getRange(i, 2).getValue(),
      });
    }
  }
  return memberList;
}
function createAndActiveSheet(member) {
  const header = [["date","subscriberCount","videoCount","viewCount","commentCount"]];
  if (!ss.getSheetByName(member.name)) {
    ss.insertSheet(member.name);
    sheet = ss.getSheetByName(member.name);
    sheet.getRange(1, 1, 1, header[0].length).setValues(header);
  }
  sheet = ss.getSheetByName(member.name);
}
function getChData(member) {
  let results = YouTube.Channels.list("snippet,statistics", {
    id: [member.id],
  });
  let tmpStaticData = results.items[0].statistics;
  let MemberDataSet = {
    commentCount: tmpStaticData.commentCount,
    videoCount: tmpStaticData.videoCount,
    subscriberCount: tmpStaticData.subscriberCount,
    viewCount: tmpStaticData.viewCount,
  };
  return MemberDataSet;
}
function insertRow(MemberDataSet) {
  const dateTime = getNowYMD();
  const lastRow = sheet.getLastRow();
  const values = [
    [
      dateTime,
      MemberDataSet.subscriberCount,
      MemberDataSet.videoCount,
      MemberDataSet.viewCount,
      MemberDataSet.commentCount,
    ],
  ];
  let dataCheck = sheet.getRange(lastRow + 1, 1).getValue() === dateTime;
  if (!dataCheck) {
    sheet.insertRows(lastRow + 1, 1);
    sheet.getRange(lastRow + 1, 1, 1, values[0].length).setValues(values);
  }
}
function getNowYMD() {
  var dt = new Date();
  var y = dt.getFullYear();
  var m = ("00" + (dt.getMonth() + 1)).slice(-2);
  var d = ("00" + dt.getDate()).slice(-2);
  var result = y + "-" + m + "-" + d;
  return result;
}
