# SharePoint document indexing with Azure AI Search

This sample continues from Step 4 of the Microsoft Learn
[SharePoint in Microsoft 365 indexer](https://learn.microsoft.com/en-us/azure/search/search-how-to-index-sharepoint-online#index-document-metadata)
procedure. Complete the article's prerequisites and Steps 1 through 3 first,
then use this project instead of manually performing Step 4 and the remaining
REST operations.

It creates or updates:

1. A SharePoint data source using a managed identity federated credential.
2. A chunk-oriented hybrid search index containing text and vector fields.
3. A skillset that splits documents and creates Azure OpenAI embeddings.
4. An indexer that extracts `contentAndMetadata` and projects one search
   document per chunk.
5. Indexer status, manual run, and hybrid query operations.

The data source only permits the document-library container modes
`defaultSiteLibrary`, `allSiteLibraries`, and `useQuery`. It doesn't use
`allSiteLists`, `allSitePages`, or `allSiteContent`.

## End-to-end setup order

Follow these phases in order. The SharePoint Learn article covers ingestion;
the Foundry steps add the embedding capability required for hybrid
keyword-plus-vector search.

### Phase 1: Complete the SharePoint prerequisites through Step 3

Use
[SharePoint in Microsoft 365 indexer](https://learn.microsoft.com/en-us/azure/search/search-how-to-index-sharepoint-online#index-document-metadata)
as the reference document.

Complete all prerequisites and the configuration that appears before
**Step 4: Create data source**:

1. Register for the SharePoint indexer preview, if required for the search
   service.
2. **Step 1:** Enable the Azure AI Search service's system-assigned managed
   identity. This is optional in the article for some authentication patterns,
   but it is required by this project's default secretless SharePoint and
   Foundry configuration.
3. **Step 2:** Choose application permissions for document-library indexing.
4. **Step 3:** Create the Microsoft Entra ingestion app.
5. Add Microsoft Graph application permissions:
   - `Files.Read.All`
   - `Sites.Read.All`
6. Grant tenant admin consent.
7. Configure the ingestion app's federated credential to trust the Azure AI
   Search managed identity, following the article's secretless authentication
   instructions.

Stop when you reach **Step 4: Create data source**. Don't manually create the
data source, index, or indexer from the article. The Python `provision` command
creates those resources and additionally creates the vector index and
skillset.

### Phase 2: Create the Foundry embedding deployment

If you don't already have a suitable Microsoft Foundry/Azure OpenAI resource:

1. Create a Microsoft Foundry parent resource and project.
2. In the model catalog, select a supported embedding model. This project
   defaults to `text-embedding-3-small`.
3. Deploy the model. A clear deployment name is
   `sharepoint-embeddings`.
4. From the parent resource's **Keys and Endpoint** page, copy its supported
   model endpoint. It commonly resembles one of these values:

   ```text
   https://<resource-name>.services.ai.azure.com
   https://<resource-name>.openai.azure.com
   ```

   Use the parent resource endpoint, not the Foundry project browser URL or
   project endpoint. In particular, remove a path such as
   `/api/projects/<project-name>`:

   ```text
   Correct:   https://<resource-name>.services.ai.azure.com
   Incorrect: https://<resource-name>.services.ai.azure.com/api/projects/<project-name>
   ```
5. Record the deployment name, model name, and dimensions for `.env`:

   ```text
   AZURE_OPENAI_ENDPOINT=https://<resource-name>.services.ai.azure.com
   AZURE_OPENAI_EMBEDDING_DEPLOYMENT=sharepoint-embeddings
   AZURE_OPENAI_EMBEDDING_MODEL=text-embedding-3-small
   AZURE_OPENAI_EMBEDDING_DIMENSIONS=1536
   ```

Supported embedding models are `text-embedding-ada-002`,
`text-embedding-3-small`, and `text-embedding-3-large`.

### Phase 3: Grant Azure AI Search access to Foundry

Grant the search service managed identity permission to call the embedding
deployment. This is an Azure RBAC role assignment on the **parent
Foundry/Azure OpenAI resource**.

The required direction is:

```text
Azure AI Search managed identity
    → calls the Foundry-hosted embedding deployment
```

In the Azure portal:

1. Open the parent Foundry/Azure OpenAI resource, not just the Foundry project.
2. Select **Access control (IAM)**.
3. Select **Add** → **Add role assignment**.
4. Select the **Cognitive Services OpenAI User** role.
5. For **Assign access to**, select **Managed identity**.
6. Select the managed identity belonging to the Azure AI Search service.
7. Review and assign the role.

Don't create the connection shown under the Foundry project's **Manage** →
**Connected resources** → **Azure AI Search** for this purpose. That connection
configures the reverse direction:

```text
Foundry project managed identity
    → calls Azure AI Search
```

It can be useful for a Foundry agent, but it isn't required by this indexing
script. Copilot Studio connects to the completed Azure AI Search index
separately as a knowledge source.

### Phase 4: Prepare the machine running this project

- Install Python 3.10 or later.
- Install Azure CLI and sign in with `az login`.
- Give the signed-in identity the **Search Service Contributor** role on the
  Azure AI Search service.
- Give the signed-in identity **Search Index Data Reader** if you will run the
  sample `search` commands. Alternatively, use `AZURE_SEARCH_ADMIN_KEY` for
  both management and query operations.

The signed-in identity authorizes this Python script to manage Azure AI Search
resources. It is separate from the search service managed identity used for
SharePoint and Foundry outbound connections.

## Phase 5: Configure the local environment

Create and activate a virtual environment, then install dependencies:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Copy `.env.example` to `.env`, replace the SharePoint, Search, and Azure OpenAI
placeholders, and load it into the current PowerShell session:

```powershell
Copy-Item .env.example .env
Get-Content .env |
  Where-Object { $_ -match '^\s*[^#].*=.*$' } |
  ForEach-Object {
    $name, $value = $_ -split '=', 2
    Set-Item -Path "Env:$($name.Trim())" -Value $value.Trim()
  }
```

### Keep configuration safe in GitHub

Commit `.env.example` as the public configuration reference, but keep only
placeholder or fictitious values in it. Store actual resource names, tenant and
application IDs, endpoints, identity resource IDs, and keys in `.env`.

The repository's `.gitignore` ignores `.env` and all `.env.*` variants, while
explicitly allowing `.env.example`. Before committing, verify the result:

```powershell
git check-ignore -v .env
git check-ignore -v .env.example
git status --short
```

The first command should report that `.env` is ignored. The second should
either produce no output or show the negation rule, and `git status` should
list `.env.example` but never `.env`.

Environment variables such as tenant IDs and resource endpoints aren't
passwords, but they disclose details about your Azure environment. Keep the
public example generic. Never put an admin key, client secret, access token, or
real credential in `.env.example`. For deployed applications, use managed
identity or a secret store such as Azure Key Vault instead of checking secrets
into source control.

For system-assigned managed identity, don't set
`USER_ASSIGNED_IDENTITY_RESOURCE_ID`. For user-assigned managed identity, set
it to the full Azure resource ID. In both cases,
`FEDERATED_CREDENTIAL_APPLICATION_ID` is the managed identity's application
(client) ID, not its object/principal ID and not `ENTRA_APPLICATION_ID`.

### Identify the correct federated credential ID

For a system-assigned setup, `FEDERATED_CREDENTIAL_APPLICATION_ID` must contain
the **application (client) ID of the Azure AI Search service's managed
identity**.

The GUID displayed as **Subject identifier or claims matching expression** on
the ingestion app's **Federated credentials** page is the managed identity's
**object/principal ID**. That subject value proves which identity is trusted,
but it isn't the value required by the data source's
`FederatedCredentialApplicationId` connection-string property.

Use the Azure AI Search identity's principal ID to look up its corresponding
service principal, then return the service principal's `appId`:

```powershell
$principalId = az search service show `
  --name <search-service-name> `
  --resource-group <resource-group-name> `
  --query identity.principalId `
  --output tsv

az ad sp show `
  --id $principalId `
  --query appId `
  --output tsv
```

Put the output of the second command into:

```text
FEDERATED_CREDENTIAL_APPLICATION_ID=<application-client-id-from-appId>
```

For a user-assigned managed identity, retrieve its client ID directly:

```powershell
az identity show `
  --name <managed-identity-name> `
  --resource-group <resource-group-name> `
  --query clientId `
  --output tsv
```

ID comparison:

| Value | Where it comes from | Environment variable |
|---|---|---|
| SharePoint ingestion app client ID | App registration **Overview** → Application (client) ID | `ENTRA_APPLICATION_ID` |
| Managed identity application/client ID | The commands above | `FEDERATED_CREDENTIAL_APPLICATION_ID` |
| Managed identity object/principal ID | Search service **Identity** page; also shown as the federated credential subject | Don't put this in the connection string |
| Federated credential name | For example, `searchToSPCred` | Not used by the REST payload |

The default `allSiteLibraries` mode indexes document files from every document
library in the configured site. To select one library or folder, use:

```powershell
$env:SP_CONTAINER_NAME = "useQuery"
$env:SP_CONTAINER_QUERY = "includeLibrary=https://contoso.sharepoint.com/sites/MySite/Shared Documents"
```

## Phase 6: Execute the REST provisioning code

Open PowerShell in this project folder. Activate the virtual environment and
make sure the dependencies are installed:

```powershell
.\.venv\Scripts\Activate.ps1
python -m pip install -r .\requirements.txt
```

The script automatically loads `.env` from the current project directory by
using `python-dotenv`. You don't need to export each value into PowerShell.

If you want the values available to other commands in the same PowerShell
session, you can optionally load them manually:

```powershell
Get-Content .\.env |
  Where-Object { $_ -match '^\s*[^#].*=.*$' } |
  ForEach-Object {
    $name, $value = $_ -split '=', 2
    Set-Item -Path "Env:$($name.Trim())" -Value $value.Trim()
  }
```

Run this preflight check before sending requests to Azure. It checks that all
required variables exist and that common template placeholders haven't been
left in `.env`:

```powershell
$requiredVariables = @(
  "AZURE_SEARCH_SERVICE",
  "SHAREPOINT_SITE_URL",
  "ENTRA_APPLICATION_ID",
  "ENTRA_TENANT_ID",
  "FEDERATED_CREDENTIAL_APPLICATION_ID",
  "AZURE_OPENAI_ENDPOINT",
  "AZURE_OPENAI_EMBEDDING_DEPLOYMENT"
)

$invalidVariables = foreach ($name in $requiredVariables) {
  $value = [Environment]::GetEnvironmentVariable($name)
  if (
    [string]::IsNullOrWhiteSpace($value) -or
    $value -eq "my-search-service" -or
    $value -match "contoso\.sharepoint\.com" -or
    $value -match "my-openai-resource" -or
    $value -eq "00000000-0000-0000-0000-000000000000"
  ) {
    $name
  }
}

if ($invalidVariables) {
  throw "Set real values for: $($invalidVariables -join ', ')"
}

Write-Host "Environment preflight passed."
```

Sign in to Azure for `DefaultAzureCredential`. Use an identity that has
**Search Service Contributor** on the Azure AI Search service. To execute the
sample search queries, it also needs **Search Index Data Reader**:

```powershell
az login --use-device-code
az account show --output table
```

Create or update the resources in this order:

1. SharePoint data source
2. Hybrid text-and-vector index
3. Text Split and Azure OpenAI embedding skillset
4. SharePoint indexer

The indexer starts its initial ingestion run when it is created. The following
command waits up to one hour for that run:

```powershell
python .\sharepoint_search_setup.py provision --wait --timeout 3600
```

The `provision` command uses idempotent REST `PUT` operations, so it can be run
again after a configuration change. A successful run ends with an indexer
result whose `status` is `success`.

If the command times out or you want to inspect indexing details, check status:

```powershell
python .\sharepoint_search_setup.py status
```

After correcting any reported permission or configuration error, start another
indexer run and wait for completion:

```powershell
python .\sharepoint_search_setup.py run --wait --timeout 3600
```

Confirm that SharePoint chunks were indexed:

```powershell
python .\sharepoint_search_setup.py search "*"
```

If this command returns HTTP 403, the identity used by `az login` can manage
the search service but can't read index documents. Assign that signed-in user
the **Search Index Data Reader** role. Run these commands manually:

```powershell
$searchServiceName = Get-Content .\.env |
  Where-Object { $_ -match '^AZURE_SEARCH_SERVICE=' } |
  ForEach-Object { ($_ -split '=', 2)[1].Trim() }

if ([string]::IsNullOrWhiteSpace($searchServiceName)) {
  throw "AZURE_SEARCH_SERVICE is missing from .env"
}

$searchServiceId = az resource list `
  --name $searchServiceName `
  --resource-type Microsoft.Search/searchServices `
  --query "[0].id" `
  --output tsv

if ([string]::IsNullOrWhiteSpace($searchServiceId)) {
  throw "Search service '$searchServiceName' wasn't found in the active Azure subscription"
}

$signedInUserId = az ad signed-in-user show `
  --query id `
  --output tsv

if ([string]::IsNullOrWhiteSpace($signedInUserId)) {
  throw "The signed-in Azure CLI account couldn't be resolved as a user"
}

az role assignment create `
  --assignee-object-id $signedInUserId `
  --assignee-principal-type User `
  --role "Search Index Data Reader" `
  --scope $searchServiceId
```

Role assignments can take several minutes to propagate. Wait and rerun:

```powershell
python .\sharepoint_search_setup.py search "*" --top 3
```

Creating the role assignment requires permission to manage IAM, such as Owner
or User Access Administrator. If you don't have that permission, ask the Azure
resource owner to assign the role.

Then run a hybrid keyword-plus-vector search:

```powershell
python .\sharepoint_search_setup.py search "annual report" --top 20
```

After provisioning, the Azure portal should show these resources under the
search service:

```text
Data sources  → sharepoint-datasource
Indexes       → sharepoint-hybrid-index
Skillsets     → sharepoint-hybrid-skillset
Indexers      → sharepoint-indexer
```

For normal text, `search` sends the same text as the lexical `search` clause and
as a `kind: text` vector query against `content_vector`. Azure AI Search
vectorizes the query with the index's Azure OpenAI vectorizer and fuses keyword
and vector rankings. A literal `*` is a browse operation and doesn't invoke the
vectorizer.

The hybrid schema is chunk based. If the earlier nonvector version has already
populated `sharepoint-index`, use the new default name
`sharepoint-hybrid-index` (recommended) or delete and recreate the old index
before provisioning. Updating an existing populated index won't remove its
old parent-shaped documents.

The script defaults to the `2026-08-01-preview` REST API from the Learn article.
Preview API versions can change, so override `AZURE_SEARCH_API_VERSION` if the
article is updated.
