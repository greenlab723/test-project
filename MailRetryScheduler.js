function installMailRetryTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const hasTrigger = triggers.some((trigger) => trigger.getHandlerFunction() === 'processMailRetryQueue');
  if (hasTrigger) {
    return;
  }
  ScriptApp.newTrigger('processMailRetryQueue').timeBased().everyHours(1).create();
}

function removeMailRetryTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((trigger) => {
    if (trigger.getHandlerFunction() === 'processMailRetryQueue') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}
