export const msalConfig = {
  auth: {
    clientId: 'a8e08cda-65ee-4686-97ba-e7d9e0f8a63f',
    authority: 'https://login.microsoftonline.com/d43d7b87-367a-4e2c-9e40-9ded6a42bf83',
    redirectUri: 'https://tthschool.github.io/weekly-product/redirect',
    postLogoutRedirectUri: '/',
  },
  cache: {
    cacheLocation: 'localStorage',
    storeAuthStateInCookie: false,
  },
};

export const loginRequest = {
  scopes: ["User.Read", "Mail.Read"]
};


export function buildWeeklyMailUri() {
  const now = new Date();
  const day = now.getDay();
  const diffToMonday = (day === 0 ? -6 : 1 - day);
  const start = new Date(now);
  start.setDate(now.getDate() + diffToMonday);
  start.setHours(0, 0, 0, 0);

  const end = new Date(start);
  end.setDate(start.getDate() + 6);
  end.setHours(23, 59, 59, 999);

  const startIso = start.toISOString();
  const endIso = end.toISOString();

  return `https://graph.microsoft.com/v1.0/me/messages?$search="subject:週報"&$filter=receivedDateTime ge ${startIso} and receivedDateTime le ${endIso}`;
}
