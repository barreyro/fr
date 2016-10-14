function buildRosterOptions() {
  var rosterFiles = [];
  var rosterNames = [];
  var userName = DriveRoster.getEffectiveUserName();
  try {
    rosterFiles = DriveRoster.getAvailableRosterFiles();
    var rosters = DriveRoster.getRosters(rosterFiles);
  } catch(err) {
    return [];
  }
 
  var options = [];
  var j = 0;
  for (var i=0; i<rosters.length; i++) {
    var thisClassName = rosters[i].className;
    var thisOwner = rosters[i].owner;
    if (thisOwner === userName) {
      options.push({name: "My Roster - " + thisClassName + " - First Last", rangeId: "roster||" + thisClassName + "||" + thisOwner + "||firstlast"});
      options.push({name:  "My Roster - " + thisClassName + " - Last, First", rangeId: "roster||" + thisClassName + "||" + thisOwner + "||lastfirst"});
    } else {
      options.push({name: "Shared Roster - " + thisClassName + " - First Last", rangeId: "roster||" + thisClassName + "||" + thisOwner + "||firstlast"});
      options.push({name:  "Shared Roster - " + thisClassName + " - Last, First", rangeId: "roster||" + thisClassName + "||" + thisOwner + "||lastfirst"});
    }
  }
  return options;
}


function buildRosterList(rangeId) {
  var list = [];
  var rosterName = rangeId.split("||")[1];
  var owner = rangeId.split("||")[2];
  var listType = rangeId.split("||")[3];
  var roster = DriveRoster.getRosterByName("roster||" + owner + "||" + rosterName);
  if (roster) {
    var students = roster.students;
    for (var i=0; i<students.length; i++) {
      if (listType === "firstlast") {
        list.push(students[i].firstName + " " + students[i].lastName);
      }
      if (listType === "lastfirst") {
        list.push(students[i].lastName + ", " + students[i].firstName);
      }
    }
  }
  return list;
}


function testBuildRoster() {
  var options = buildRosterOptions();
  var rosterList = buildRosterList(options[0].rangeId);
  debugger;
}
