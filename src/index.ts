import { AgentSession } from './session';
import type { Env } from './types';

export { AgentSession };

export default {
  async fetch(request: Request, env: Env): Promise<Response> {
    const { pathname } = new URL(request.url);

    if (pathname === '/ws') {
      const id = env.SESSION.newUniqueId();
      return env.SESSION.get(id).fetch(request);
    }

    return env.ASSETS.fetch(request);
  },
} satisfies ExportedHandler<Env>;
