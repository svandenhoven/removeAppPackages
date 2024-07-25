const { DeviceCodeCredential } = require('@azure/identity');
const { Client, ResponseType } = require('@microsoft/microsoft-graph-client');
const readline = require('readline');

require('isomorphic-fetch'); // Required for @azure/identity

var deleteAnswer = 'y';

// Function to ask for user confirmation
function askForConfirmation(question) {
    if (deleteAnswer.toLowerCase() === 'a') {
        return Promise.resolve(true);
    }

    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout
    });

    return new Promise((resolve) => {
        rl.question(question, (answer) => {
            rl.close();
            deleteAnswer = answer;
            resolve(answer.trim().toLowerCase() === 'y');
        });
    });
}

async function getAuthenticatedClient(clientId, tenantId) {
	const credential = new DeviceCodeCredential({
		tenantId: tenantId,
		clientId: clientId,
		userPromptCallback: (info) => console.log(info.message) // This will log the device code message to the console
	});

	const client = Client.initWithMiddleware({
		debugLogging: true,
		authProvider: {
			getAccessToken: async () => {
				try {
					const tokenResponse = await credential.getToken("https://graph.microsoft.com/.default");
					return tokenResponse.token;
				} catch (error) {
					throw error;
				}
			}
		}
	});

	return client;
}

async function main() {
    const args = process.argv.slice(2); // Remove the first two elements
    if (args.length !== 3 && args.length !== 4) {
        console.error('Please provide the app package name, ClientId and TenantId as an argument. Version is optional.');
        console.error('Usage: node index.js <appPackageName> <ClientId> <TenantId> (<Version>)');
        return;
    }

    // Get the app package name from the arguments
    const name = args[0];
    const clientId = args[1];
    const tenantId = args[2];
    const version = args[3] ?? '';

	const client = await getAuthenticatedClient(clientId, tenantId);
	const user = await client.api('/me').get();
	console.log(user);

    const url = `/appCatalogs/teamsApps?$expand=appDefinitions($select=id, version, publishingState)&$filter=distributionMethod eq 'organization' ${version !== '' ? 'and appDefinitions/any(a:a/version ne \'' + version + '\')' : '' }and displayName eq '${name}'`;
    const appPackages = await client.api(url).get();
    for (const appPackage of appPackages.value) {
        console.log('-----------------------------------');
        console.log(`displayName: ${appPackage.displayName}`);
        console.log(`id: ${appPackage.id}`);
        if (appPackage.appDefinitions.length > 0) { 
            console.log(`version: ${appPackage.appDefinitions[0].version}`);
            console.log(`definitionId: ${appPackage.appDefinitions[0].id}`);
            console.log(`publishingState: ${appPackage.appDefinitions[0].publishingState}`);

            // Ask the user for confirmation to delete
            const isConfirmed = await askForConfirmation('Do you want to delete this app package? (Y/N/A(ll)): ');
            if (isConfirmed) {
                var deleteUrl = '';
                switch (appPackage.appDefinitions[0].publishingState) {
                    case 'submitted':
                    case 'rejected':
                        deleteUrl = `/appCatalogs/teamsApps/${appPackage.id}/appDefinitions/${appPackage.appDefinitions[0].id}`;
                        break;
                    case 'published':
                        deleteUrl = `/appCatalogs/teamsApps/${appPackage.id}`;
                        break;
                    default:
                        deleteUrl = `/appCatalogs/teamsApps/${appPackage.id}`;
                        break;
                }
                try {
                    await client.api(deleteUrl).delete();
                    console.log(`App package ${appPackage.displayName} deleted successfully`);
                } catch (error) {
                    console.error(`Error deleting app package ${appPackage.displayName}: ${error}`);
                }
            } else {
                console.log('Deletion cancelled by the user.');
            }
        }
        console.log('-----------------------------------');
    }
}

main().catch(console.error);