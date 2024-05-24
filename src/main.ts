import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';

import { AppModule } from './app/app.module';

Office.initialize = () => {
  const platform = platformBrowserDynamic();
  platformBrowserDynamic().bootstrapModule(AppModule)
  .catch(err => console.error(err));
}
