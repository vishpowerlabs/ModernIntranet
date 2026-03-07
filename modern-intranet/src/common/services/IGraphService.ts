/**
 * DEVELOPER BY VISHPOWERLABS
 * CONTACT : INFO@VISHPOWERLABS.COM
 */

export interface IGraphService {
    graphGet<T>(endpoint: string): Promise<T>;
    graphPost<T>(endpoint: string, body: object): Promise<T>;
}
