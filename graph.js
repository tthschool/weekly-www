export const getGraphClient = (accessToken) => {
  return MicrosoftGraph.Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    }
  });
};
