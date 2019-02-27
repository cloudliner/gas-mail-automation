function setupTrigger() {
  // delete all triggers
  const allTriggers = ScriptApp.getProjectTriggers();
  for (const trigger of allTriggers) {
    ScriptApp.deleteTrigger(trigger);
  }

  ScriptApp.newTrigger("automaticArchive")
    .timeBased()
    .everyMinutes(5)
    .create();

  ScriptApp.newTrigger("automaticLabel")
    .timeBased()
    .everyMinutes(5)
    .create();

  ScriptApp.newTrigger("automaticDelete")
    .timeBased()
    .atHour(4)
    .everyDays(1)
    .inTimezone("Asia/Tokyo")
    .create();

  return "success";
}
