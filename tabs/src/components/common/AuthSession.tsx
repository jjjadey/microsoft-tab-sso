// If no request to cdn./rap, cookies will be expired around 2 hours.
// Keep session alive by request to session.html.

import axios, { AxiosRequestConfig } from 'axios';

const REQ_TIMER = 1000 * 60 * 90; // 90 minutes
let timeoutId: NodeJS.Timeout;
let isStart = false;

const startAuthSession = () => {
  if (isStart) {
    return;
  }
  isStart = true;

  localStorage.setItem('teamsdirapp_nextAuthRefresh', String(Date.now() + REQ_TIMER));
  timerRequestSession();
};

async function requestSession() {
  const nextRefreshTime = Number(localStorage.getItem('teamsdirapp_nextAuthRefresh'));

  const nowDate = Date.now();
  if ((nextRefreshTime - 30000) <= nowDate) {
    localStorage.setItem('teamsdirapp_nextAuthRefresh', String(nowDate + REQ_TIMER));
    try {
      const axiosConfig: AxiosRequestConfig = {
        url: './session.html',
      };

      const response = await axios.request(axiosConfig);
      console.info('Request session status:', response.status);
    } catch (error) {
      console.info('Request session failed:', error);
    }
  }

  timerRequestSession();
}

function timerRequestSession() {
  if (timeoutId) {
    clearTimeout(timeoutId);
  }

  timeoutId = setTimeout(() => {
    requestSession();
  }, REQ_TIMER);
}

export default startAuthSession;
