function getLabelObject(labelName: string): GoogleAppsScript.Gmail.GmailLabel | null {
  let labelObject = null;
  if (labelName && labelName.trim().length !== 0) {
    labelName = labelName.trim();
    labelObject = GmailApp.getUserLabelByName(labelName) ?
      GmailApp.getUserLabelByName(labelName) : GmailApp.createLabel(labelName);
  }
  return labelObject;
}

function getLabelObjectList(labelNames: string): GoogleAppsScript.Gmail.GmailLabel[] {
  const labelNameList = labelNames ? labelNames.split(",") : [];
  const labelList = [] as GoogleAppsScript.Gmail.GmailLabel[];
  labelNameList.forEach((labelName) => {
    const labelObject = getLabelObject(labelName);
    if (labelObject) {
      labelList.push(labelObject);
    }
  });
  return labelList;
}

function getLabelNames(labels: GoogleAppsScript.Gmail.GmailLabel[]) {
  let labelNames = "";
  labels.forEach((label) => {
    labelNames = labelNames + label.getName() + ", ";
  });
  return labelNames;
}
