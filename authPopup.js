
import { buildWeeklyMailUri,msalConfig, loginRequest } from "./authConfig.js";
import { callGraph } from "./fetch.js";

const myMSALObj = new msal.PublicClientApplication(msalConfig);
let username = "";

function selectAccount() {
  const currentAccounts = myMSALObj.getAllAccounts();
  if (currentAccounts.length > 0) {
    username = currentAccounts[0].username;
  }
}

export function signIn() {
  myMSALObj.loginPopup(loginRequest).then(handleResponse).catch(console.error);
}

export function signOut() {
  const account = myMSALObj.getAccountByUsername(username);
  myMSALObj.logoutPopup({ account }).catch(console.error);
}

function handleResponse(response) {
  if (response !== null) {
    username = response.account.username;
  } else {
    selectAccount();
  }
}

export async function getMail() {
  const uri = buildWeeklyMailUri();
  return await callGraph(
    username,
    ["User.Read", "Mail.Read"],
    uri,
    msal.InteractionType.Popup,
    myMSALObj
  );
}

selectAccount();
window.signIn = signIn;
window.signOut = signOut;
