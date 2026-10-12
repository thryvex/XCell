import React from 'react';
import { th } from '../i18n/th';
import { APP_INFO } from '../appInfo';

export const App: React.FC = () => {
  return (
    <div style={{ padding: 20, display: 'flex', flexDirection: 'column', height: '100%' }}>
      <h1>{APP_INFO.name}</h1>
      <p>{th.statusBar.ready}</p>
    </div>
  );
};

export default App;
