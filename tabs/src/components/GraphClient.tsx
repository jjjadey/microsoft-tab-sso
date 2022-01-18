import { useEffect, useState } from "react";
import * as microsoftTeams from "@microsoft/teams-js";

const GraphClient = () => {
  const [myProfile, setMyProfile] = useState<any>();
  const [myClientSideToken, setMyClientSideToken] = useState<any>("");

  useEffect(() => {
    const getToken = async () => {
      try {
        const clientSideToken = await getClientSideToken();
        setMyClientSideToken(clientSideToken);
        console.log(">>>clientSideToken", clientSideToken);
        const userProfile = await getUserProfile(clientSideToken);
        console.log(">>userProfile", userProfile);
        setMyProfile(userProfile);
      } catch (error) {
        console.error(error);
      }
    };
    getToken();
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Get a client side token from Teams
  const getClientSideToken = async () => {
    microsoftTeams.initialize();
    return new Promise((resolve, reject) => {
      microsoftTeams.authentication.getAuthToken({
        successCallback: (result) => {
          resolve(result);
        },
        failureCallback: (error) => {
          reject(error);
        },
      });
    });
  };

  // Get the user profile from our web service
  const getUserProfile = async (clientSideToken: any) => {
    if (!clientSideToken) {
      // eslint-disable-next-line no-throw-literal
      throw "Error: No client side token provided in getUserProfile()";
    }

    // Get Teams context, converting callback to a promise so we can await it
    const context: any = await (() => {
      return new Promise((resolve) => {
        microsoftTeams.getContext((context) => resolve(context));
      });
    })();

    // Request the user profile from our web service
    let serverUrl = `${process.env.REACT_APP_TEAMSFX_ENDPOINT}/userProfile`;

    const response = await fetch(serverUrl, {
      method: "post",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        tenantId: context.tid,
        clientSideToken: clientSideToken,
      }),
      cache: "default",
    });

    if (response.ok) {
      const userProfile = await response.json();
      return userProfile;
    } else {
      const error = await response.json();
      console.error(error);
      showConsentPopup();
      // client display a pop-up auth box

      throw error;
    }
  };

  // Display the consent pop-up if needed
  const showConsentPopup = async () => {
    await microsoftTeams.authentication.authenticate({
      //   url: window.location.origin + "/auth-start.html",
      url:
        process.env.REACT_APP_START_LOGIN_PAGE_URL ||
        "https://localhost:53000/auth-start.html",
      width: 600,
      height: 535,
      successCallback: () => {
        console.log("Got success callback");
      },
    });
  };

  return (
    <div>
      <h1>My client side token</h1>
      <p>{myClientSideToken}</p>
      <h1>My profile</h1>
      <p>name: {myProfile?.displayName}</p>
      <p>mail: {myProfile?.mail}</p>
      <p>id: {myProfile?.id}</p>
    </div>
  );
};

export default GraphClient;