export interface Env {
  SESSION: DurableObjectNamespace;
  SHEETS: R2Bucket;
  OPENAI_API_KEY: string;
  ASSETS: Fetcher;
}
