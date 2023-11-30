import { useEffect, useState } from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import { Button } from "@fluentui/react-northstar";

const GraphClient = () => {
  const [myProfile, setMyProfile] = useState<any>();
  const [myPresence, setMyPresence] = useState<any>();
  const [myClientSideToken, setMyClientSideToken] = useState<any>("");

  useEffect(() => {
    const getToken = async () => {
      try {
        const clientSideToken = await getClientSideToken();
        setMyClientSideToken(clientSideToken);
        console.log(">>>clientSideToken", clientSideToken);
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
      setMyProfile(userProfile);
    } else {
      const error = await response.json();
      console.error(error);
      showConsentPopup();
      // client display a pop-up auth box

      throw error;
    }
  };

  const getPresence = async (clientSideToken: any) => {
    if (!clientSideToken) {
      // eslint-disable-next-line no-throw-literal
      throw "Error: No client side token provided in getPresence()";
    }

    // Get Teams context, converting callback to a promise so we can await it
    const context: any = await (() => {
      return new Promise((resolve) => {
        microsoftTeams.getContext((context) => resolve(context));
      });
    })();

    // Request the user profile from our web service
    let serverUrl = `${process.env.REACT_APP_TEAMSFX_ENDPOINT}/presence`;

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
      const res = await response.json();
      setMyPresence(res)
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
    let scope = process.env.REACT_APP_SCOPE?.split(",");
    scope = scope?.map((s) => `https://graph.microsoft.com/${s}`);

    let authUrl: string = process.env.REACT_APP_START_LOGIN_PAGE_URL || "";

    authUrl += `?clicentId=${process.env.REACT_APP_CLIENT_ID}`;
    authUrl += `&scope=${encodeURIComponent(scope?.join(" ") || "")}`;

    console.log(">>>authUrl", authUrl);

    console.log(">>> showConsentPopup");

    await microsoftTeams.authentication.authenticate({
      //   url: window.location.origin + "/auth-start.html",
      url: authUrl,
      width: 600,
      height: 535,
      successCallback: () => {
        console.log("Got success callback");
        // window.location.reload();
      },
      failureCallback: (err) => {
        console.error(err);
      },
    });
  };

  return (
    <div>
      <h1>My client side token</h1>
      <p>{myClientSideToken}</p>
      <h1>My profile</h1>
      <Button onClick={() => getUserProfile(myClientSideToken)} primary>
        get profile
      </Button>
      {myProfile && (
        <div>
          <p>name: {myProfile.displayName}</p>
          <p>mail: {myProfile.mail}</p>
          <p>id: {myProfile.id}</p>
        </div>
      )}
      <h1>My presence</h1>
      <Button onClick={() => getPresence(myClientSideToken)} primary>
        get presence
      </Button>
      {myPresence && (
        <div>
          <p>activity: {myPresence.activity}</p>
          <p>availability: {myPresence.availability}</p>
          <p>id: {myPresence.id}</p>
        </div>
      )}
    </div>
  );
};

export default GraphClient;