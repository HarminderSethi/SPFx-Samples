import type { MSGraphClientFactory, MSGraphClientV3 } from '@microsoft/sp-http';

/**
 * Singleton wrapper around `MSGraphClientV3`. Initialise once during
 * the command set's `onInit()` and reuse everywhere via `getInstance()`.
 */
export class GraphService {
  private static _instance: GraphService | undefined;
  private readonly _client: MSGraphClientV3;

  private constructor(client: MSGraphClientV3) {
    this._client = client;
  }

  public static async initialize(factory: MSGraphClientFactory): Promise<GraphService> {
    if (!GraphService._instance) {
      const client = await factory.getClient('3');
      GraphService._instance = new GraphService(client);
    }
    return GraphService._instance;
  }

  public static getInstance(): GraphService {
    if (!GraphService._instance) {
      throw new Error('GraphService not initialized. Call GraphService.initialize() first.');
    }
    return GraphService._instance;
  }

  public static reset(): void {
    GraphService._instance = undefined;
  }

  public async get<T>(endpoint: string): Promise<T> {
    try {
      return (await this._client.api(endpoint).get()) as T;
    } catch (error) {
      throw GraphService._normalizeError(error, `GET ${endpoint}`);
    }
  }

  public async post<T>(endpoint: string, body: unknown): Promise<T> {
    try {
      return (await this._client.api(endpoint).post(body)) as T;
    } catch (error) {
      throw GraphService._normalizeError(error, `POST ${endpoint}`);
    }
  }

  private static _normalizeError(error: unknown, context: string): Error {
    if (error instanceof Error) {
      const message = error.message || 'Unknown Graph error';
      return new Error(`[GraphService] ${context} failed: ${message}`);
    }
    return new Error(`[GraphService] ${context} failed.`);
  }
}
