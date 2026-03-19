import { DurableObject } from 'cloudflare:workers';
import { TOOLS, executeTool } from './tools';
import type { Env } from './types';

const REALTIME_URL = 'wss://api.openai.com/v1/realtime?model=gpt-4o-realtime-preview-2024-12-17';

const SESSION_CONFIG = {
  modalities: ['text', 'audio'],
  instructions:
    'You are a helpful spreadsheet assistant with full access to an Excel file. ' +
    'When asked about the spreadsheet, call get_sheet_info first, then read the relevant cells or ranges. ' +
    'When making changes: read first, make the edit, then confirm what changed with specific values. ' +
    'Keep responses concise and conversational.',
  voice: 'alloy',
  input_audio_format: 'pcm16',
  output_audio_format: 'pcm16',
  input_audio_transcription: { model: 'whisper-1' },
  turn_detection: {
    type: 'server_vad',
    threshold: 0.5,
    prefix_padding_ms: 300,
    silence_duration_ms: 800,
  },
  tools: TOOLS,
  tool_choice: 'auto',
};

export class AgentSession extends DurableObject<Env> {
  async fetch(request: Request): Promise<Response> {
    if (request.headers.get('Upgrade') !== 'websocket') {
      return new Response('Expected WebSocket', { status: 400 });
    }

    // ── Inbound WebSocket (browser) ───────────────────────────────────────────
    const [client, server] = Object.values(new WebSocketPair());
    server.accept();

    // ── Outbound WebSocket (OpenAI Realtime API) ──────────────────────────────
    let openai: WebSocket;
    try {
      const resp = await fetch(REALTIME_URL, {
        headers: {
          Authorization: `Bearer ${this.env.OPENAI_API_KEY}`,
          'OpenAI-Beta': 'realtime=v1',
        },
      } as RequestInit);
      // @ts-ignore — Workers exposes .webSocket on upgraded responses
      openai = resp.webSocket;
      if (!openai) throw new Error('No WebSocket in response');
      openai.accept();
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      server.send(JSON.stringify({ type: 'error', message: `OpenAI connection failed: ${msg}` }));
      server.close(1011, 'OpenAI connection failed');
      return new Response(null, { status: 101, webSocket: client });
    }

    // Configure the Realtime session immediately
    openai.send(JSON.stringify({ type: 'session.update', session: SESSION_CONFIG }));

    // ── OpenAI → browser ──────────────────────────────────────────────────────
    openai.addEventListener('message', async (evt: MessageEvent) => {
      let msg: Record<string, unknown>;
      try {
        msg = JSON.parse(evt.data as string);
      } catch {
        return;
      }

      // Intercept tool calls — execute server-side, never expose to browser raw
      if (msg.type === 'response.function_call_arguments.done') {
        let result: unknown;
        try {
          const toolArgs = JSON.parse(msg.arguments as string) as Record<string, unknown>;
          result = await executeTool(msg.name as string, toolArgs, this.env.SHEETS);
        } catch (err: unknown) {
          result = { error: err instanceof Error ? err.message : String(err) };
        }

        openai.send(JSON.stringify({
          type: 'conversation.item.create',
          item: { type: 'function_call_output', call_id: msg.call_id, output: JSON.stringify(result) },
        }));
        openai.send(JSON.stringify({ type: 'response.create' }));

        // Notify the browser so it can show the activity sidebar
        try {
          server.send(JSON.stringify({ type: 'tool_activity', name: msg.name, result }));
        } catch { /* browser may have disconnected */ }
        return;
      }

      // Forward everything else to the browser
      try { server.send(evt.data as string); } catch { /* ignore */ }
    });

    openai.addEventListener('close', () => { try { server.close(); } catch { /* ignore */ } });
    openai.addEventListener('error', () => { try { server.close(1011, 'OpenAI error'); } catch { /* ignore */ } });

    // ── Browser → OpenAI ──────────────────────────────────────────────────────
    server.addEventListener('message', (evt: MessageEvent) => {
      try { openai.send(evt.data as string); } catch { /* ignore */ }
    });
    server.addEventListener('close', () => { try { openai.close(); } catch { /* ignore */ } });

    return new Response(null, { status: 101, webSocket: client });
  }
}
