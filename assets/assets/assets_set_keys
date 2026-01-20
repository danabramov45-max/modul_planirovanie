// ####################################################################################################################################
// OZON УСТАНОВКА И УДАЛЕНИЕ КЛЮЧЕЙ API
// ####################################################################################################################################

function ozonSetSellerAccess(service) {
  var ui = SpreadsheetApp.getUi();
  const cient_id = ui.prompt("Введите Client ID (Seller API)", ui.ButtonSet.OK);

  if (cient_id.getSelectedButton() != ui.Button.OK) {
    ss.toast("Отмена", "⚠️ ATTENTION", 10);
    return;
  }

  const api_key = ui.prompt("Введите API-ключ (Seller API)", ui.ButtonSet.OK);

  if (api_key.getSelectedButton() != ui.Button.OK) {
    ss.toast("Отмена", "⚠️ ATTENTION", 10);
    return;
  }

  if (!api_key.getResponseText()) {
    ss.toast("Нет данных", "⚠️ ATTENTION", 10);
    return;
  }

  const clientId = cient_id.getResponseText().trim();
  const apiKey = api_key.getResponseText().trim();

  const storage =
    context?.storage ||
    (service && setStorage({ service, serviceType: context?.serviceType })) ||
    null;

  storage.refresh({ clientId, apiKey });

  ss.toast("Сохранили Client ID и API-ключ", "✅ SUCCESS", 10);
}

function ozonSetPerformanceAccess(service) {
  var ui = SpreadsheetApp.getUi();
  let client_id = ui.prompt(
    "Введите Client ID (Performance API)",
    ui.ButtonSet.OK
  );

  if (client_id.getSelectedButton() != ui.Button.OK) {
    ss.toast("Отмена", "⚠️ ATTENTION", 10);
    return;
  }

  if (!client_id.getResponseText()) {
    ss.toast("Нет данных", "⚠️ ATTENTION", 10);
    return;
  }

  client_id = client_id.getResponseText().trim();

  let client_secret = ui.prompt(
    "Введите Client Secret (Performance API)",
    ui.ButtonSet.OK
  );

  if (client_secret.getSelectedButton() != ui.Button.OK) {
    ss.toast("Отмена", "⚠️ ATTENTION", 10);
    return;
  }

  if (!client_secret.getResponseText()) {
    ss.toast("Нет данных", "⚠️ ATTENTION", 10);
    return;
  }

  client_secret = client_secret.getResponseText().trim();

  try {
    if (validateClientSecret(client_secret) && validateClientId(client_id)) {
      const storage =
        context?.storage ||
        (service &&
          setStorage({ service, serviceType: context?.serviceType })) ||
        null;

      storage.refresh({ client_id, client_secret });

      ss.toast("Сохранили Client ID и Client Secret", "✅ SUCCESS", 10);
    }
  } catch (e) {
    console.log(error.stack);
    ui.alert(`⚠️ ATTENTION`, error.stack, ui.ButtonSet.OK);
  }
}

function ozonDeleteSellerAccess(service) {
  const storage =
    context?.storage ||
    (service && setStorage({ service, serviceType: context?.serviceType })) ||
    null;
  storage.clear();
  ss.toast("Удалили Client ID и API-ключ", "✅ SUCCESS", 10);
}

function ozonSetFolder(service) {
  var ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "Вставьте ссылку на папку или ее ID",
    "Папка, куда будут сохраняться этикетки.\nID в строке браузера: https://drive.google.com/drive/folders/*******это id*******",
    ui.ButtonSet.OK
  );

  if (response.getSelectedButton() != ui.Button.OK) {
    ss.toast("Отмена", "⚠️ ATTENTION", 10);
    return;
  }

  if (!response.getResponseText()) {
    ss.toast("Ссылка не добавлена", "⚠️ ATTENTION", 10);
    return;
  }

  const folderId = response
    .getResponseText()
    .trim()
    .match(/[-\w]{25,}/);

  if (!folderId) {
    ss.toast("Не верная ссылка!!!", "⚠️ ATTENTION", 10);
    return;
  }

  const storage =
    context?.storage ||
    (service && setStorage({ service, serviceType: context?.serviceType })) ||
    null;

  storage.refresh({ folderId: folderId[0] });

  ss.toast("Папка добавлена!!!", "✅ SUCCESS", 10);
}

// ####################################################################################################################################
// МОЙ СКЛАД УСТАНОВКА И УДАЛЕНИЕ КЛЮЧА API
// ####################################################################################################################################

function moyskladSetToken(service) {
  var ui = SpreadsheetApp.getUi();

  const token = ui.prompt(
    `Введите Токен МойСклад`,
    "Генерируется в ЛК:\nНастройки > Токены",
    ui.ButtonSet.OK
  );

  if (token.getSelectedButton() != ui.Button.OK) {
    ss.toast("Отмена", "⚠️ ATTENTION", 10);
    return;
  }

  if (!token.getResponseText()) {
    ss.toast("Нет данных", "⚠️ ATTENTION", 10);
    return;
  }

  const storage = new StorageService({
    service: context?.service || service,
    serviceType: context?.serviceType,
    serviceName: "Moysklad",
  });

  storage.refresh({ token: token.getResponseText().trim() });
  ss.toast("Сохранили Токен МойСклад", "✅ SUCCESS", 10);
}

