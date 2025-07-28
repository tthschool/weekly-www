import { getGraphClient } from "./graph.js";

export const callGraph = async (username, scopes, uri, interactionType, myMSALObj) => {
  const account = myMSALObj.getAccountByUsername(username);
  let tokenResponse;

  try {
    tokenResponse = await myMSALObj.acquireTokenSilent({ account, scopes });
  } catch {
    if (interactionType === msal.InteractionType.Popup) {
      tokenResponse = await myMSALObj.acquireTokenPopup({ scopes });
    } else {
      throw new Error("Token acquisition failed");
    }
  }

  const client = getGraphClient(tokenResponse.accessToken);
  const response = await client.api(uri).get();
  return response;
};
