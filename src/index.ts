/**
 * Azure Functions entry point.
 * Side-effecting handler registrations.
 */

import 'dotenv/config';
import './handlers/webhook';
import './handlers/renew';
import './handlers/summary';
import './handlers/statements';
import './handlers/rentals';
import './handlers/reconcile';
