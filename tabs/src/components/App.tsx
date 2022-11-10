import React
  // , { useState, useEffect }
  from 'react';
// import * as microsoftTeams from '@microsoft/teams-js';
import { HashRouter as Router, Route } from 'react-router-dom';
import { Provider, Loader, teamsTheme } from '@fluentui/react-northstar';
// import { Provider as RTProvider, themeNames } from '@fluentui/react-teams';

import { useTeamsFx } from '../utils/useTeamsFx';

// components
import Signin from './Signin';

// functions
import startAuthSession from './common/AuthSession';

/**
 * The main app which handles the initialization and routing
 * of the app.
 */

const App: React.FC = () => {
  const { theme, loading } = useTeamsFx();
  startAuthSession();
  // console.info('theme', theme);

  // // const [appContext, setAppContext] = useState<microsoftTeams.Context>();
  // const [appAppearance, setAppAppearance] = useState<themeNames>(
  //   themeNames.Default,
  // );

  // useEffect(() => {
  //   /**
  //    * With the context properties in hand, your app has a solid understanding
  //    * of what's happening around it in Teams.
  //    * https://docs.microsoft.com/en-us/javascript/api/@microsoft/teams-js/context?view=msteams-client-js-latest&preserve-view=true
  //    * */
  //   microsoftTeams.getContext((context) => {
  //     // setAppContext(context);
  //     setAppAppearance(initTeamsTheme(context.theme));

  //     /**
  //      * Tells Microsoft Teams platform that we are done saving our settings.
  // Microsoft Teams waits
  //      * for the app to call this API before it dismisses the dialog. If the wait times out,
  //      * you will
  //      * see an error indicating that the configuration settings could not be saved.
  //      * */
  //     microsoftTeams.appInitialization.notifySuccess();
  //   });

  //   /**
  //    * Theme change handler
  //    * https://docs.microsoft.com/en-us/javascript/api/@microsoft/teams-js/microsoftteams?view=msteams-client-js-latest#registerOnThemeChangeHandler__theme__string_____void_
  //    * */
  //   microsoftTeams.registerOnThemeChangeHandler((theme) => {
  //     setAppAppearance(initTeamsTheme(theme));
  //   });
  // }, []);

  return (
    // <RTProvider themeName={appAppearance} lang="en-US">
    <Provider
      theme={theme || teamsTheme}
    >
      <Provider>
        <Router>
          {loading ? (
            <Loader style={{ margin: 100 }} />
          ) : (
            <>
              <Route exact path="/signin" component={Signin} />
            </>
          )}
        </Router>
      </Provider>
    </Provider>
    // </RTProvider>
  );
};

export default App;

// Possible values for theme: 'default', 'light', 'dark' and 'contrast'
// function initTeamsTheme(theme: string | undefined): themeNames {
//   switch (theme) {
//     case 'dark':
//       return themeNames.Dark;
//     case 'contrast':
//       return themeNames.HighContrast;
//     default:
//       return themeNames.Default;
//   }
// }
