import { MSGraphClientV3 } from '@microsoft/sp-http';
import { IGraphService } from './IGraphService';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export class GraphService implements IGraphService {
    private readonly _context: WebPartContext;
    private _client: MSGraphClientV3 | null = null;

    constructor(context: WebPartContext) {
        this._context = context;
    }

    private async getClient(): Promise<MSGraphClientV3> {
        if (this._client) {
            return this._client;
        }
        this._client = await this._context.msGraphClientFactory.getClient('3');
        return this._client;
    }

    public async graphGet<T>(endpoint: string): Promise<T> {
        try {
            const client = await this.getClient();
            const response = await client.api(endpoint).get();
            return response as T;
        } catch (error) {
            console.error(`GraphService.graphGet error for endpoint ${endpoint}:`, error);
            throw error;
        }
    }

    public async graphPost<T>(endpoint: string, body: object): Promise<T> {
        try {
            const client = await this.getClient();
            const response = await client.api(endpoint).post(body);
            return response as T;
        } catch (error) {
            console.error(`GraphService.graphPost error for endpoint ${endpoint}:`, error);
            throw error;
        }
    }
}
