/**
 * Repository Mapper
 * 
 * Maps SharePoint Embedded containers to CMIS repositories.
 */

import type { FileStorageContainer, Drive } from '@microsoft/microsoft-graph-types';
import { RepositoryInfo, RepositoryCapabilities, GetRepositoriesResponse } from '../types/cmis';

/**
 * Default CMIS capabilities for SPE-backed repositories.
 * These reflect what SharePoint Embedded supports through our adapter.
 * Per CMIS 1.1 spec section 2.1.1.
 */
export const DEFAULT_CAPABILITIES: RepositoryCapabilities = {
    capabilityContentStreamUpdatability: 'anytime',
    capabilityChanges: 'none',                    // Change log not supported in MVP
    capabilityRenditions: 'none',                 // Renditions not supported in MVP
    capabilityGetDescendants: true,
    capabilityGetFolderTree: true,
    capabilityMultifiling: false,                 // SPE doesn't support multi-filing
    capabilityUnfiling: false,                    // SPE doesn't support unfiling
    capabilityVersionSpecificFiling: false,
    capabilityPWCSearchable: false,
    capabilityPWCUpdatable: true,
    capabilityAllVersionsSearchable: false,
    capabilityOrderBy: 'common',
    capabilityQuery: 'metadataonly',              // Simple name-based query only (see cmis/query.ts)
    capabilityJoin: 'none',
    capabilityACL: 'manage',                      // Can both read and manage ACLs (see cmis/aclMapper.ts)
};

/**
 * Product information constants
 */
export const PRODUCT_INFO = {
    vendorName: 'Microsoft',
    productName: 'SharePoint Embedded CMIS Adapter',
    productVersion: '1.0.0',
    cmisVersionSupported: '1.1',
};

/**
 * Maps an SPE container to a CMIS RepositoryInfo object.
 * 
 * @param container - The SharePoint Embedded container
 * @param drive - The container's drive (optional, used to get root folder ID)
 * @param baseUrl - Base URL for constructing rootFolderUrl
 */
export function mapContainerToRepository(
    container: FileStorageContainer,
    drive?: Drive,
    baseUrl?: string
): RepositoryInfo {
    // The root folder ID is either from the drive's root, or we construct it
    // In SPE, the root is typically the drive's root item ID
    const rootFolderId = drive?.root?.id || 'root';
    const repositoryId = container.id!;

    return {
        repositoryId,
        repositoryName: container.displayName!,
        repositoryDescription: container.description || '',
        vendorName: PRODUCT_INFO.vendorName,
        productName: PRODUCT_INFO.productName,
        productVersion: PRODUCT_INFO.productVersion,
        rootFolderId,
        rootFolderUrl: baseUrl ? `${baseUrl}/${repositoryId}/root` : undefined,
        repositoryUrl: baseUrl ? `${baseUrl}/${repositoryId}` : undefined,
        latestChangeLogToken: null,
        cmisVersionSupported: PRODUCT_INFO.cmisVersionSupported,
        thinClientURI: null,
        changesIncomplete: true,
        changesOnType: [],
        principalIdAnonymous: 'anonymous',
        principalIdAnyone: 'anyone',
        capabilities: DEFAULT_CAPABILITIES,
    };
}

/**
 * Maps multiple SPE containers to a CMIS GetRepositoriesResponse.
 * Per CMIS 1.1 Browser Binding spec section 5.2.1, returns repositories
 * directly at the root level keyed by repositoryId.
 * 
 * @param containers - Array of SharePoint Embedded containers
 * @param baseUrl - Base URL for constructing rootFolderUrl
 */
export function mapContainersToRepositories(
    containers: FileStorageContainer[],
    baseUrl?: string
): GetRepositoriesResponse {
    const repositories: GetRepositoriesResponse = {};

    for (const container of containers) {
        const repository = mapContainerToRepository(container, undefined, baseUrl);
        repositories[repository.repositoryId] = repository;
    }

    return repositories;
}

/**
 * Filters containers by name prefix (for filter parameter support)
 * 
 * @param containers - Array of containers to filter
 * @param filter - Filter string (prefix match on displayName)
 */
export function filterContainersByName(
    containers: FileStorageContainer[],
    filter?: string
): FileStorageContainer[] {
    if (!filter) {
        return containers;
    }

    const lowerFilter = filter.toLowerCase();
    return containers.filter(c => 
        c.displayName?.toLowerCase().startsWith(lowerFilter)
    );
}
