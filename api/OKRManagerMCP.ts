import {McpServer} from '@modelcontextprotocol/sdk/server/mcp.js'
import {z} from 'zod'
import {DurableMCP} from "./lib/MCPEntrypoint.ts";
import {okrService} from "./OKRService.ts";
import {AuthenticationContext, Objective} from "../types";


/**
 * The `OKRManagerMCP` class exposes the OKR Manager Service via the Model Context Protocol
 * for consumption by API Agents
 */
export class OKRManagerMCP extends DurableMCP<AuthenticationContext, Env> {
    async init() {
    }

    get okrService() {
        console.log('Binding service to tenant', this.props.claims['https://stytch.com/organization'].organization_id);
        return okrService(this.env, this.props.claims['https://stytch.com/organization'].organization_id)
    }

    // withRequiredScope checks that the authentication context contains the scope requested for the specific function
    // TODO: This should function at the Permission level, not the Scope level
    // HACKY!
    withRequiredScope = <T extends CallableFunction>(scope: string, fn: T): T => {
        const withRequiredScopeImpl = (...args: unknown[]) => {
            console.log(this.props.claims, scope, 'xxxx')
            if (!this.props.claims.scope.split(' ').includes(scope)) {
                throw new Error(`Caller does not have access to required scope ${scope}`)
            }
            return fn(...args)
        }
        return withRequiredScopeImpl as unknown as T
    }

    formatResponse = (description: string, newState: Objective[]): {
        content: Array<{ type: 'text', text: string }>
    } => {
        return {
            content: [{
                type: "text",
                text: `Success! ${description}\n\nNew state:\n${JSON.stringify(newState, null, 2)}`
            }]
        };
    }

    get server() {
        const server = new McpServer({
            name: 'TODO Service',
            version: '1.0.0',
        })

        // server.resource("Todos", new ResourceTemplate("todoapp://todos/{id}", {
        //         list: async () => {
        //             const todos = await this.todoService.get()
        //
        //             return {
        //                 resources: todos.map(todo => ({
        //                     name: todo.text,
        //                     uri: `todoapp://todos/${todo.id}`
        //                 }))
        //             }
        //         }
        //     }),
        //     async (uri, {id}) => {
        //         const todos = await this.todoService.get();
        //         const todo = todos.find(todo => todo.id === id);
        //         return {
        //             contents: [
        //                 {
        //                     uri: uri.href,
        //                     text: todo ? `text: ${todo.text} completed: ${todo.completed}` : 'NOT FOUND',
        //                 },
        //             ],
        //         }
        //     },
        // )

        server.tool('listObjectives', 'View all objectives and key results for the organization', this.withRequiredScope('read:okrs', async () => {
            const result = await this.okrService.get()
            return this.formatResponse('Objectives retrieved successfully', result);
        }))

        server.tool('addObjective', 'Add a new top-level objective for the organization', {objectiveText: z.string()}, this.withRequiredScope('manage:okrs', async ({objectiveText}) => {
            const result = await this.okrService.addObjective(objectiveText)
            return this.formatResponse('Objective added successfully', result);
        }))

        server.tool('deleteObjective', 'Remove an existing top-level objective from the organization', {okrID: z.string()}, this.withRequiredScope('manage:okrs', async ({okrID}) => {
            const result = await this.okrService.deleteObjective(okrID);
            return this.formatResponse('Objective deleted successfully', result);
        }));

        server.tool('addKeyResult', 'Add a new key result to a specific objective', {
            okrID: z.string(),
            keyResultText: z.string()
        }, this.withRequiredScope('manage:krs', async ({okrID, keyResultText}) => {
            const result = await this.okrService.addKeyResult(okrID, keyResultText);
            return this.formatResponse('Key result added successfully', result);
        }));

        server.tool('setKeyResultAttainment', 'Set the attainment value for a specific key result in a specific objective', {
            okrID: z.string(),
            keyResultID: z.string(),
            attainment: z.number().int().min(0).max(100)
        }, this.withRequiredScope('report_kr_status', async ({okrID, keyResultID, attainment}) => {
            const result = await this.okrService.setKeyResultAttainment(okrID, keyResultID, attainment);
            return this.formatResponse('Key result attainment set successfully', result);
        }));

        server.tool('deleteKeyResult', 'Remove a key result from a specific objective', {
            okrID: z.string(),
            keyResultID: z.string()
        }, this.withRequiredScope('manage:krs', async ({okrID, keyResultID}) => {
            const result = await this.okrService.deleteKeyResult(okrID, keyResultID);
            return this.formatResponse('Key result deleted successfully', result);
        }));

        return server
    }
}