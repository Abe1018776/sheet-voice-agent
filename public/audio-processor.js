// AudioWorklet processor: resamples mic input to 24kHz PCM16 for OpenAI Realtime API
class PCMProcessor extends AudioWorkletProcessor {
  constructor() {
    super();
    this.buffer = [];
    this.TARGET_RATE = 24000;
    this.CHUNK_SAMPLES = 2400; // 100ms at 24kHz
  }

  process(inputs) {
    const input = inputs[0];
    if (!input || !input[0]) return true;

    const samples = input[0];
    const ratio = this.TARGET_RATE / sampleRate; // sampleRate is global in AudioWorklet
    const resampledLen = Math.floor(samples.length * ratio);

    // Linear resampling
    for (let i = 0; i < resampledLen; i++) {
      const srcIdx = Math.min(Math.floor(i / ratio), samples.length - 1);
      this.buffer.push(samples[srcIdx]);
    }

    // Emit chunks
    while (this.buffer.length >= this.CHUNK_SAMPLES) {
      const chunk = this.buffer.splice(0, this.CHUNK_SAMPLES);
      const pcm16 = new Int16Array(chunk.length);
      for (let i = 0; i < chunk.length; i++) {
        pcm16[i] = Math.max(-32768, Math.min(32767, Math.round(chunk[i] * 32767)));
      }
      this.port.postMessage(pcm16.buffer, [pcm16.buffer]);
    }

    return true;
  }
}

registerProcessor('pcm-processor', PCMProcessor);
