function getLabelObject(labelName:string) {
  var labelObject = null;
  if (labelName && labelName.trim().length != 0) {
    labelName = labelName.trim();
    labelObject = GmailApp.getUserLabelByName(labelName) ? GmailApp.getUserLabelByName(labelName) : GmailApp.createLabel(labelName);
  }
  return labelObject;
}

function getLabelObjectList(labelNames:string) {
  var labelNameList = labelNames ? labelNames.split(',') : [];
  var labelList = [];
  labelNameList.forEach(function(labelName) {
    var labelObject = getLabelObject(labelName);
    labelList.push(labelObject);
  });
  return labelList;
}

function getLabelNames(labels) {
  var labelNames = '';
  labels.forEach(function(label) {
    labelNames = labelNames + label.getName() + ', ';
  });
  return labelNames;
}
