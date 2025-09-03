// Expose application version at runtime
// Webpack supports importing JSON; this reads version from package.json at build time
import pkg from '../../../package.json';

export const APP_VERSION = pkg?.version || '0.0.0';
