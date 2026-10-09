from __future__ import annotations

import argparse
import json
import os
import sys
import time
from dataclasses import dataclass
from typing import Any
from urllib.parse import urlparse

import requests
from azure.identity import DefaultAzureCredential
from dotenv import load_dotenv


DEFAULT_API_VERSION = "2026-08-01-preview"
DOCUMENT_LIBRARY_CONTAINERS = {
    "defaultSiteLibrary",
    "allSiteLibraries",
    "useQuery",
}
DEFAULT_DOCUMENT_EXTENSIONS = (
    ".pdf,.doc,.docx,.ppt,.pptx,.xls,.xlsx,.txt,.csv,.rtf,.html,.htm"
)
SUPPORTED_EMBEDDING_DIMENSIONS = {
    "text-embedding-ada-002": (1536, 1536),
    "text-embedding-3-small": (1, 1536),
    "text-embedding-3-large": (1, 3072),
}


def required_env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise ValueError(f"Required environment variable {name} is not set.")
    return value


@dataclass(frozen=True)
class Settings:
    search_service: str
    sharepoint_site_url: str
    entra_application_id: str
    federated_credential_application_id: str
    tenant_id: str
    data_source_name: str
    index_name: str
    indexer_name: str
    skillset_name: str
    container_name: str
    container_query: str | None
    included_extensions: str | None
    api_version: str
    schedule_interval: str | None
    search_admin_key: str | None
    user_assigned_identity_resource_id: str | None
    azure_openai_endpoint: str
    azure_openai_deployment: str
    azure_openai_model: str
    embedding_dimensions: int
    azure_openai_auth_identity_resource_id: str | None
    chunk_length: int
    chunk_overlap: int

    @classmethod
    def from_environment(cls) -> "Settings":
        container_name = os.getenv("SP_CONTAINER_NAME", "allSiteLibraries").strip()
        if container_name not in DOCUMENT_LIBRARY_CONTAINERS:
            allowed = ", ".join(sorted(DOCUMENT_LIBRARY_CONTAINERS))
            raise ValueError(
                f"SP_CONTAINER_NAME must be a document-library mode: {allowed}."
            )

        container_query = os.getenv("SP_CONTAINER_QUERY", "").strip() or None
        if container_name == "useQuery" and not container_query:
            raise ValueError(
                "SP_CONTAINER_QUERY is required when SP_CONTAINER_NAME=useQuery."
            )

        extensions = os.getenv(
            "INDEXED_FILE_EXTENSIONS", DEFAULT_DOCUMENT_EXTENSIONS
        ).strip()
        model = os.getenv(
            "AZURE_OPENAI_EMBEDDING_MODEL", "text-embedding-3-small"
        ).strip()
        if model not in SUPPORTED_EMBEDDING_DIMENSIONS:
            supported = ", ".join(SUPPORTED_EMBEDDING_DIMENSIONS)
            raise ValueError(
                f"AZURE_OPENAI_EMBEDDING_MODEL must be one of: {supported}."
            )
        dimensions = int(os.getenv("AZURE_OPENAI_EMBEDDING_DIMENSIONS", "1536"))
        minimum, maximum = SUPPORTED_EMBEDDING_DIMENSIONS[model]
        if not minimum <= dimensions <= maximum:
            raise ValueError(
                f"AZURE_OPENAI_EMBEDDING_DIMENSIONS for {model} must be "
                f"between {minimum} and {maximum}."
            )
        chunk_length = int(os.getenv("CHUNK_LENGTH", "2000"))
        chunk_overlap = int(os.getenv("CHUNK_OVERLAP", "500"))
        if chunk_length <= 0:
            raise ValueError("CHUNK_LENGTH must be greater than zero.")
        if not 0 <= chunk_overlap < chunk_length:
            raise ValueError(
                "CHUNK_OVERLAP must be nonnegative and less than CHUNK_LENGTH."
            )
        azure_openai_endpoint = required_env("AZURE_OPENAI_ENDPOINT").rstrip("/")
        parsed_endpoint = urlparse(azure_openai_endpoint)
        supported_endpoint_suffixes = (
            ".openai.azure.com",
            ".services.ai.azure.com",
            ".cognitiveservices.azure.com",
        )
        if (
            parsed_endpoint.scheme != "https"
            or not parsed_endpoint.hostname
            or not parsed_endpoint.hostname.endswith(supported_endpoint_suffixes)
        ):
            raise ValueError(
                "AZURE_OPENAI_ENDPOINT must be the HTTPS endpoint of the parent "
                "Foundry/Azure OpenAI resource."
            )
        if "/api/projects/" in parsed_endpoint.path.lower():
            raise ValueError(
                "AZURE_OPENAI_ENDPOINT must not be a Foundry project endpoint. "
                "Remove the '/api/projects/<project-name>' path and use the "
                "parent resource endpoint."
            )

        return cls(
            search_service=required_env("AZURE_SEARCH_SERVICE"),
            sharepoint_site_url=required_env("SHAREPOINT_SITE_URL").rstrip("/"),
            entra_application_id=required_env("ENTRA_APPLICATION_ID"),
            federated_credential_application_id=required_env(
                "FEDERATED_CREDENTIAL_APPLICATION_ID"
            ),
            tenant_id=required_env("ENTRA_TENANT_ID"),
            data_source_name=os.getenv(
                "AZURE_SEARCH_DATASOURCE_NAME", "sharepoint-datasource"
            ).strip(),
            index_name=os.getenv(
                "AZURE_SEARCH_INDEX_NAME", "sharepoint-hybrid-index"
            ).strip(),
            indexer_name=os.getenv(
                "AZURE_SEARCH_INDEXER_NAME", "sharepoint-indexer"
            ).strip(),
            skillset_name=os.getenv(
                "AZURE_SEARCH_SKILLSET_NAME", "sharepoint-hybrid-skillset"
            ).strip(),
            container_name=container_name,
            container_query=container_query,
            included_extensions=extensions or None,
            api_version=os.getenv(
                "AZURE_SEARCH_API_VERSION", DEFAULT_API_VERSION
            ).strip(),
            schedule_interval=os.getenv("INDEXER_SCHEDULE_INTERVAL", "").strip()
            or None,
            search_admin_key=os.getenv("AZURE_SEARCH_ADMIN_KEY", "").strip()
            or None,
            user_assigned_identity_resource_id=os.getenv(
                "USER_ASSIGNED_IDENTITY_RESOURCE_ID", ""
            ).strip()
            or None,
            azure_openai_endpoint=azure_openai_endpoint,
            azure_openai_deployment=required_env(
                "AZURE_OPENAI_EMBEDDING_DEPLOYMENT"
            ),
            azure_openai_model=model,
            embedding_dimensions=dimensions,
            azure_openai_auth_identity_resource_id=os.getenv(
                "AZURE_OPENAI_AUTH_IDENTITY_RESOURCE_ID", ""
            ).strip()
            or None,
            chunk_length=chunk_length,
            chunk_overlap=chunk_overlap,
        )


class SearchRestClient:
    def __init__(self, settings: Settings) -> None:
        self.settings = settings
        self.base_url = (
            f"https://{settings.search_service}.search.windows.net"
        )
        self.session = requests.Session()
        self.credential = (
            None if settings.search_admin_key else DefaultAzureCredential()
        )

    def _headers(self) -> dict[str, str]:
        headers = {"Content-Type": "application/json"}
        if self.settings.search_admin_key:
            headers["api-key"] = self.settings.search_admin_key
        else:
            assert self.credential is not None
            token = self.credential.get_token("https://search.azure.com/.default")
            headers["Authorization"] = f"Bearer {token.token}"
        return headers

    def request(
        self,
        method: str,
        path: str,
        *,
        body: dict[str, Any] | None = None,
        params: dict[str, Any] | None = None,
        expected_statuses: tuple[int, ...] = (200,),
    ) -> Any:
        query = {"api-version": self.settings.api_version}
        if params:
            query.update(params)

        response = self.session.request(
            method,
            f"{self.base_url}{path}",
            headers=self._headers(),
            params=query,
            json=body,
            timeout=120,
        )
        if response.status_code not in expected_statuses:
            try:
                detail = json.dumps(response.json(), indent=2)
            except ValueError:
                detail = response.text
            if (
                response.status_code == 403
                and "/docs/search" in path
                and not self.settings.search_admin_key
            ):
                detail = (
                    "The signed-in identity is not authorized to query index "
                    "documents. Assign it the 'Search Index Data Reader' role "
                    "on the Azure AI Search service, wait for RBAC propagation, "
                    "and retry.\n"
                    f"{detail}"
                )
            raise RuntimeError(
                f"{method} {response.url} failed with HTTP "
                f"{response.status_code}:\n{detail}"
            )
        if not response.content:
            return None
        return response.json()

    def put(self, resource: str, name: str, body: dict[str, Any]) -> Any:
        return self.request(
            "PUT",
            f"/{resource}/{name}",
            body=body,
            expected_statuses=(200, 201, 204),
        )


def build_connection_string(settings: Settings) -> str:
    return (
        f"SharePointOnlineEndpoint={settings.sharepoint_site_url};"
        f"ApplicationId={settings.entra_application_id};"
        f"FederatedCredentialApplicationId="
        f"{settings.federated_credential_application_id};"
        f"TenantId={settings.tenant_id}"
    )


def data_source_definition(settings: Settings) -> dict[str, Any]:
    definition: dict[str, Any] = {
        "name": settings.data_source_name,
        "type": "sharepoint",
        "credentials": {"connectionString": build_connection_string(settings)},
        "container": {
            "name": settings.container_name,
            "query": settings.container_query,
        },
    }
    if settings.user_assigned_identity_resource_id:
        definition["identity"] = {
            "@odata.type": "#Microsoft.Azure.Search.DataUserAssignedIdentity",
            "userAssignedIdentity": settings.user_assigned_identity_resource_id,
        }
    return definition


def index_definition(settings: Settings) -> dict[str, Any]:
    return {
        "name": settings.index_name,
        "fields": [
            {
                "name": "id",
                "type": "Edm.String",
                "key": True,
                "searchable": True,
                "filterable": True,
                "analyzer": "keyword",
            },
            {
                "name": "parent_id",
                "type": "Edm.String",
                "searchable": False,
                "filterable": True,
            },
            {
                "name": "metadata_spo_item_name",
                "type": "Edm.String",
                "searchable": True,
                "filterable": True,
                "sortable": True,
                "facetable": False,
            },
            {
                "name": "metadata_spo_item_path",
                "type": "Edm.String",
                "searchable": True,
                "filterable": False,
                "sortable": False,
                "facetable": False,
            },
            {
                "name": "metadata_spo_item_weburi",
                "type": "Edm.String",
                "searchable": False,
                "filterable": False,
                "sortable": False,
                "facetable": False,
            },
            {
                "name": "metadata_spo_item_content_type",
                "type": "Edm.String",
                "searchable": False,
                "filterable": True,
                "sortable": False,
                "facetable": True,
            },
            {
                "name": "metadata_spo_item_extension",
                "type": "Edm.String",
                "searchable": False,
                "filterable": True,
                "sortable": False,
                "facetable": True,
            },
            {
                "name": "metadata_spo_item_last_modified",
                "type": "Edm.DateTimeOffset",
                "searchable": False,
                "filterable": True,
                "sortable": True,
                "facetable": False,
            },
            {
                "name": "metadata_spo_item_size",
                "type": "Edm.Int64",
                "searchable": False,
                "filterable": True,
                "sortable": True,
                "facetable": False,
            },
            {
                "name": "metadata_spo_site_url",
                "type": "Edm.String",
                "searchable": False,
                "filterable": True,
                "sortable": False,
                "facetable": False,
            },
            {
                "name": "content",
                "type": "Edm.String",
                "searchable": True,
                "filterable": False,
                "sortable": False,
                "facetable": False,
            },
            {
                "name": "content_vector",
                "type": "Collection(Edm.Single)",
                "searchable": True,
                "retrievable": False,
                "stored": False,
                "dimensions": settings.embedding_dimensions,
                "vectorSearchProfile": "sharepoint-vector-profile",
            },
        ],
        "vectorSearch": {
            "algorithms": [
                {
                    "name": "sharepoint-hnsw",
                    "kind": "hnsw",
                    "hnswParameters": {
                        "m": 4,
                        "efConstruction": 400,
                        "efSearch": 100,
                        "metric": "cosine",
                    },
                }
            ],
            "profiles": [
                {
                    "name": "sharepoint-vector-profile",
                    "algorithm": "sharepoint-hnsw",
                    "vectorizer": "sharepoint-openai-vectorizer",
                }
            ],
            "vectorizers": [
                {
                    "name": "sharepoint-openai-vectorizer",
                    "kind": "azureOpenAI",
                    "azureOpenAIParameters": azure_openai_parameters(settings),
                }
            ],
        },
    }


def managed_identity_definition(resource_id: str) -> dict[str, str]:
    return {
        "@odata.type": "#Microsoft.Azure.Search.DataUserAssignedIdentity",
        "userAssignedIdentity": resource_id,
    }


def azure_openai_parameters(settings: Settings) -> dict[str, Any]:
    parameters: dict[str, Any] = {
        "resourceUri": settings.azure_openai_endpoint,
        "deploymentId": settings.azure_openai_deployment,
        "modelName": settings.azure_openai_model,
    }
    if settings.azure_openai_auth_identity_resource_id:
        parameters["authIdentity"] = managed_identity_definition(
            settings.azure_openai_auth_identity_resource_id
        )
    return parameters


def skillset_definition(settings: Settings) -> dict[str, Any]:
    embedding_skill: dict[str, Any] = {
        "@odata.type": "#Microsoft.Skills.Text.AzureOpenAIEmbeddingSkill",
        "name": "embed-sharepoint-chunks",
        "context": "/document/pages/*",
        "resourceUri": settings.azure_openai_endpoint,
        "deploymentId": settings.azure_openai_deployment,
        "modelName": settings.azure_openai_model,
        "dimensions": settings.embedding_dimensions,
        "inputs": [{"name": "text", "source": "/document/pages/*"}],
        "outputs": [{"name": "embedding", "targetName": "content_vector"}],
    }
    if settings.azure_openai_auth_identity_resource_id:
        embedding_skill["authIdentity"] = managed_identity_definition(
            settings.azure_openai_auth_identity_resource_id
        )

    projected_fields = [
        "metadata_spo_item_name",
        "metadata_spo_item_path",
        "metadata_spo_item_weburi",
        "metadata_spo_item_content_type",
        "metadata_spo_item_extension",
        "metadata_spo_item_last_modified",
        "metadata_spo_item_size",
    ]
    mappings = [
        {"name": "content", "source": "/document/pages/*"},
        {
            "name": "content_vector",
            "source": "/document/pages/*/content_vector",
        },
    ]
    mappings.extend(
        {"name": field, "source": f"/document/{field}"}
        for field in projected_fields
    )

    return {
        "name": settings.skillset_name,
        "description": "Chunk and vectorize SharePoint document content.",
        "skills": [
            {
                "@odata.type": "#Microsoft.Skills.Text.SplitSkill",
                "name": "split-sharepoint-content",
                "context": "/document",
                "textSplitMode": "pages",
                "maximumPageLength": settings.chunk_length,
                "pageOverlapLength": settings.chunk_overlap,
                "maximumPagesToTake": 0,
                "unit": "characters",
                "defaultLanguageCode": "en",
                "inputs": [{"name": "text", "source": "/document/content"}],
                "outputs": [{"name": "textItems", "targetName": "pages"}],
            },
            embedding_skill,
        ],
        "indexProjections": {
            "selectors": [
                {
                    "targetIndexName": settings.index_name,
                    "parentKeyFieldName": "parent_id",
                    "sourceContext": "/document/pages/*",
                    "mappings": mappings,
                }
            ],
            "parameters": {
                "projectionMode": "skipIndexingParentDocuments",
            },
        },
    }


def indexer_definition(settings: Settings) -> dict[str, Any]:
    configuration: dict[str, Any] = {"dataToExtract": "contentAndMetadata"}
    if settings.included_extensions:
        configuration["indexedFileNameExtensions"] = settings.included_extensions

    definition: dict[str, Any] = {
        "name": settings.indexer_name,
        "dataSourceName": settings.data_source_name,
        "targetIndexName": settings.index_name,
        "skillsetName": settings.skillset_name,
        "parameters": {
            "maxFailedItems": 0,
            "maxFailedItemsPerBatch": 0,
            "configuration": configuration,
        },
    }
    if settings.schedule_interval:
        definition["schedule"] = {"interval": settings.schedule_interval}
    return definition


def provision(client: SearchRestClient, settings: Settings) -> None:
    print(f"Creating/updating data source '{settings.data_source_name}'...")
    client.put(
        "datasources",
        settings.data_source_name,
        data_source_definition(settings),
    )

    print(f"Creating/updating index '{settings.index_name}'...")
    client.put("indexes", settings.index_name, index_definition(settings))

    print(f"Creating/updating skillset '{settings.skillset_name}'...")
    client.put(
        "skillsets",
        settings.skillset_name,
        skillset_definition(settings),
    )

    print(f"Creating/updating indexer '{settings.indexer_name}'...")
    client.put(
        "indexers",
        settings.indexer_name,
        indexer_definition(settings),
    )
    print("Provisioning requests completed.")


def get_status(client: SearchRestClient, settings: Settings) -> dict[str, Any]:
    status = client.request(
        "GET",
        f"/indexers/{settings.indexer_name}/status",
    )
    print(json.dumps(status, indent=2))
    return status


def run_indexer(
    client: SearchRestClient, settings: Settings
) -> str | None:
    status = client.request(
        "GET",
        f"/indexers/{settings.indexer_name}/status",
    )
    previous_start_time = (status.get("lastResult") or {}).get("startTime")
    client.request(
        "POST",
        f"/indexers/{settings.indexer_name}/run",
        expected_statuses=(202, 204, 409),
    )
    print(
        f"Indexer '{settings.indexer_name}' was started or is already running."
    )
    return previous_start_time


def wait_for_indexer(
    client: SearchRestClient,
    settings: Settings,
    timeout_seconds: int,
    previous_start_time: str | None = None,
) -> None:
    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline:
        status = client.request(
            "GET",
            f"/indexers/{settings.indexer_name}/status",
        )
        last_result = status.get("lastResult") or {}
        state = last_result.get("status")
        start_time = last_result.get("startTime")
        print(f"Indexer status: {state or status.get('status', 'unknown')}")
        if (
            previous_start_time
            and start_time == previous_start_time
            and state in {"success", "transientFailure", "persistentFailure"}
        ):
            time.sleep(10)
            continue
        if state == "success":
            print(json.dumps(last_result, indent=2))
            return
        if state in {"transientFailure", "persistentFailure"}:
            raise RuntimeError(json.dumps(last_result, indent=2))
        time.sleep(10)
    raise TimeoutError(
        f"Indexer did not finish within {timeout_seconds} seconds."
    )


def search_documents(
    client: SearchRestClient, settings: Settings, search_text: str, top: int
) -> None:
    body: dict[str, Any] = {
        "search": search_text,
        "count": True,
        "top": top,
        "select": (
            "id,parent_id,content,metadata_spo_item_name,"
            "metadata_spo_item_path,metadata_spo_item_weburi,"
            "metadata_spo_item_extension,metadata_spo_item_last_modified"
        ),
    }
    if search_text != "*":
        body["vectorQueries"] = [
            {
                "kind": "text",
                "text": search_text,
                "fields": "content_vector",
                "k": max(top, 50),
            }
        ]

    result = client.request(
        "POST",
        f"/indexes/{settings.index_name}/docs/search",
        body=body,
    )
    print(json.dumps(result, indent=2))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Configure and operate an Azure AI Search SharePoint indexer."
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    provision_parser = subparsers.add_parser(
        "provision",
        help="Create or update the data source, hybrid index, skillset, and indexer.",
    )
    provision_parser.add_argument(
        "--wait",
        action="store_true",
        help="Wait for the indexer's initial run to finish.",
    )
    provision_parser.add_argument(
        "--timeout",
        type=int,
        default=1800,
        help="Maximum wait time in seconds (default: 1800).",
    )

    subparsers.add_parser("status", help="Get the current indexer status.")

    run_parser = subparsers.add_parser("run", help="Start the indexer.")
    run_parser.add_argument(
        "--wait", action="store_true", help="Wait for the run to finish."
    )
    run_parser.add_argument(
        "--timeout",
        type=int,
        default=1800,
        help="Maximum wait time in seconds (default: 1800).",
    )

    search_parser = subparsers.add_parser(
        "search", help="Run a hybrid keyword and vector query."
    )
    search_parser.add_argument("text", nargs="?", default="*")
    search_parser.add_argument("--top", type=int, default=10)
    return parser.parse_args()


def main() -> int:
    try:
        load_dotenv()
        args = parse_args()
        settings = Settings.from_environment()
        client = SearchRestClient(settings)

        if args.command == "provision":
            provision(client, settings)
            if args.wait:
                previous_start_time = run_indexer(client, settings)
                wait_for_indexer(
                    client,
                    settings,
                    args.timeout,
                    previous_start_time,
                )
        elif args.command == "status":
            get_status(client, settings)
        elif args.command == "run":
            previous_start_time = run_indexer(client, settings)
            if args.wait:
                wait_for_indexer(
                    client,
                    settings,
                    args.timeout,
                    previous_start_time,
                )
        elif args.command == "search":
            search_documents(client, settings, args.text, args.top)
        return 0
    except (ValueError, RuntimeError, TimeoutError, requests.RequestException) as error:
        print(f"Error: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
