import React, { useState } from "react";
import "../styles/App.css";
import * as msTeams from "@microsoft/teams-js";

function App() {
  const [userToken, setUserToken] = useState<string>("No Token Available");

  msTeams.initialize();

  msTeams.authentication.getAuthToken({
    successCallback: (token: string) => {
      setUserToken(token);
      msTeams.appInitialization.notifySuccess();
    },
    failureCallback: (message: string) => {
      msTeams.appInitialization.notifyFailure({
        reason: msTeams.appInitialization.FailedReason.AuthFailed,
        message
      });
    }
  });

  return (
      <div className="App">
        <p>Sample to Demonstrate SSO in Teams Tab.</p>
        <p>
          The value of user token is : <br /> {userToken}
        </p>        
      </div>
  );
}

export default App;
