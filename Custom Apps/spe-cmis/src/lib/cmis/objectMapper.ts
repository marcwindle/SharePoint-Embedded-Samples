/**
 * Object Mapper
 * 
 * Maps Microsoft Graph DriveItems to CMIS ObjectData format.
 * Per CMIS 1.1 Browser Binding specification.
 */

import type { DriveItem } from '@microsoft/microsoft-graph-types';
import { ObjectData, ObjectList, SuccinctProperties, PropertyData, CmisPropertyType } from '../types/cmis';

/**
 * CMIS base type IDs
 */
export const CMIS_TYPE_DOCUMENT = 'cmis:document';
export const CMIS_TYPE_FOLDER = 'cmis:folder';

/**
 * Helper to create a full PropertyData object
 */
function createProperty(
    id: string,
    type: CmisPropertyType,
    value: string | number | boolean | unknown[] | null
): PropertyData {
    return {
        id,
        type,
        cardinality: 'single',
        value,
    };
}

/**
 * Maps a DriveItem to CMIS full properties format.
 * Full format: property ID -> PropertyData object with metadata.
 */
export function mapDriveItemToProperties(item: DriveItem): Record<string, PropertyData> {
    const isFolder = !!item.folder;
    const baseTypeId = isFolder ? CMIS_TYPE_FOLDER : CMIS_TYPE_DOCUMENT;

    const props: Record<string, PropertyData> = {
        'cmis:objectId': createProperty('cmis:objectId', 'id', item.id!),
        'cmis:objectTypeId': createProperty('cmis:objectTypeId', 'id', baseTypeId),
        'cmis:baseTypeId': createProperty('cmis:baseTypeId', 'id', baseTypeId),
        'cmis:name': createProperty('cmis:name', 'string', item.name!),
        'cmis:createdBy': createProperty('cmis:createdBy', 'string', item.createdBy?.user?.displayName || 'Unknown'),
        'cmis:creationDate': createProperty('cmis:creationDate', 'datetime', item.createdDateTime ? new Date(item.createdDateTime).getTime() : null),
        'cmis:lastModifiedBy': createProperty('cmis:lastModifiedBy', 'string', item.lastModifiedBy?.user?.displayName || 'Unknown'),
        'cmis:lastModificationDate': createProperty('cmis:lastModificationDate', 'datetime', item.lastModifiedDateTime ? new Date(item.lastModifiedDateTime).getTime() : null),
        'cmis:changeToken': createProperty('cmis:changeToken', 'string', item.eTag || null),
    };

    if (isFolder) {
        // Folder-specific properties
        const path = item.parentReference?.path 
            ? `${item.parentReference.path.replace(/^\/drive\/root:?/, '')}/${item.name}`
            : `/${item.name}`;
        props['cmis:path'] = createProperty('cmis:path', 'string', path);
    } else {
        // Document-specific properties
        props['cmis:contentStreamLength'] = createProperty('cmis:contentStreamLength', 'integer', item.size || 0);
        props['cmis:contentStreamMimeType'] = createProperty('cmis:contentStreamMimeType', 'string', item.file?.mimeType || 'application/octet-stream');
        props['cmis:contentStreamFileName'] = createProperty('cmis:contentStreamFileName', 'string', item.name!);
        props['cmis:contentStreamId'] = createProperty('cmis:contentStreamId', 'id', item.id!);
        props['cmis:isLatestVersion'] = createProperty('cmis:isLatestVersion', 'boolean', true);
        props['cmis:isMajorVersion'] = createProperty('cmis:isMajorVersion', 'boolean', true);
        props['cmis:isLatestMajorVersion'] = createProperty('cmis:isLatestMajorVersion', 'boolean', true);
        props['cmis:versionLabel'] = createProperty('cmis:versionLabel', 'string', '1.0');
        props['cmis:isVersionSeriesCheckedOut'] = createProperty('cmis:isVersionSeriesCheckedOut', 'boolean', false);
    }

    // Parent folder ID
    if (item.parentReference?.id) {
        props['cmis:parentId'] = createProperty('cmis:parentId', 'id', item.parentReference.id);
    }

    return props;
}

/**
 * Maps a DriveItem to CMIS succinct properties format.
 * Succinct format: property ID -> value directly (no metadata).
 */
export function mapDriveItemToSuccinctProperties(item: DriveItem): SuccinctProperties {
    const isFolder = !!item.folder;
    const baseTypeId = isFolder ? CMIS_TYPE_FOLDER : CMIS_TYPE_DOCUMENT;

    const props: SuccinctProperties = {
        'cmis:objectId': item.id!,
        'cmis:objectTypeId': baseTypeId,
        'cmis:baseTypeId': baseTypeId,
        'cmis:name': item.name!,
        'cmis:createdBy': item.createdBy?.user?.displayName || 'Unknown',
        'cmis:creationDate': item.createdDateTime ? new Date(item.createdDateTime).getTime() : null,
        'cmis:lastModifiedBy': item.lastModifiedBy?.user?.displayName || 'Unknown',
        'cmis:lastModificationDate': item.lastModifiedDateTime ? new Date(item.lastModifiedDateTime).getTime() : null,
        'cmis:changeToken': item.eTag || null,
    };

    if (isFolder) {
        // Folder-specific properties
        props['cmis:path'] = item.parentReference?.path 
            ? `${item.parentReference.path.replace(/^\/drive\/root:?/, '')}/${item.name}`
            : `/${item.name}`;
    } else {
        // Document-specific properties
        props['cmis:contentStreamLength'] = item.size || 0;
        props['cmis:contentStreamMimeType'] = item.file?.mimeType || 'application/octet-stream';
        props['cmis:contentStreamFileName'] = item.name!;
        props['cmis:contentStreamId'] = item.id!;
        props['cmis:isLatestVersion'] = true;
        props['cmis:isMajorVersion'] = true;
        props['cmis:isLatestMajorVersion'] = true;
        props['cmis:versionLabel'] = '1.0';
        props['cmis:isVersionSeriesCheckedOut'] = false;
    }

    // Parent folder ID
    if (item.parentReference?.id) {
        props['cmis:parentId'] = item.parentReference.id;
    }

    return props;
}

/**
 * Maps a DriveItem to CMIS ObjectData format.
 * 
 * @param item - The DriveItem from Graph API
 * @param succinct - If true, include succinct properties format (default: true)
 * @param includeAllowableActions - If true, include allowable actions
 */
export function mapDriveItemToObjectData(
    item: DriveItem,
    succinct: boolean = true,
    includeAllowableActions: boolean = false
): ObjectData {
    const objectData: ObjectData = {
        // Always include full properties - cmislib needs this
        properties: mapDriveItemToProperties(item),
    };

    if (succinct) {
        objectData.succinctProperties = mapDriveItemToSuccinctProperties(item);
    }

    if (includeAllowableActions) {
        objectData.allowableActions = mapDriveItemToAllowableActions(item);
    }

    return objectData;
}

/**
 * Maps DriveItem permissions to CMIS allowable actions.
 * This is a simplified implementation - real implementation would check actual permissions.
 */
export function mapDriveItemToAllowableActions(item: DriveItem): Record<string, boolean> {
    const isFolder = !!item.folder;
    const isRoot = !item.parentReference?.id || item.parentReference.path === '/drive/root:';

    const actions: Record<string, boolean> = {
        'canGetProperties': true,
        'canGetObjectParents': !isRoot,
        'canUpdateProperties': true,
        'canDeleteObject': !isRoot,
        'canMoveObject': !isRoot,
        'canGetACL': true,
        'canApplyACL': true,
    };

    if (isFolder) {
        actions['canGetChildren'] = true;
        actions['canGetDescendants'] = true;
        actions['canGetFolderTree'] = true;
        actions['canCreateDocument'] = true;
        actions['canCreateFolder'] = true;
        actions['canDeleteTree'] = !isRoot;
        actions['canGetFolderParent'] = !isRoot;
    } else {
        actions['canGetContentStream'] = true;
        actions['canSetContentStream'] = true;
        actions['canDeleteContentStream'] = true;
        actions['canGetAllVersions'] = true;
        // Backed by Graph's checkout/checkin/discardCheckout driveItem
        // actions. We don't track checked-out state server-side (Graph's
        // base driveItem shape doesn't expose it), so these are always
        // advertised as available for documents.
        actions['canCheckOut'] = true;
        actions['canCancelCheckOut'] = true;
        actions['canCheckIn'] = true;
    }

    return actions;
}

/**
 * Maps an array of DriveItems to a CMIS ObjectList.
 * Per CMIS 1.1 Browser Binding spec, each object is wrapped in an 'object' key.
 * 
 * @param items - Array of DriveItems
 * @param hasMoreItems - Whether there are more items available
 * @param numItems - Total count of items (optional)
 * @param succinct - Use succinct properties format
 * @param includeAllowableActions - Include allowable actions
 */
export function mapDriveItemsToObjectList(
    items: DriveItem[],
    hasMoreItems: boolean = false,
    numItems?: number,
    succinct: boolean = true,
    includeAllowableActions: boolean = false
): ObjectList {
    return {
        objects: items.map(item => ({
            object: mapDriveItemToObjectData(item, succinct, includeAllowableActions)
        })),
        hasMoreItems,
        numItems,
    };
}

/**
 * Maps a newly created DriveItem to CMIS ObjectData response.
 * Used after createDocument or createFolder operations.
 */
export function mapCreatedItemToObjectData(
    item: DriveItem,
    succinct: boolean = true
): ObjectData {
    return mapDriveItemToObjectData(item, succinct, true);
}

/**
 * Maps an array of DriveItems to a CMIS query results response (cmisaction=query).
 * Per CMIS 1.1 Browser Binding spec section 5.5.1, this is a distinct shape
 * from ObjectList: results are NOT wrapped in an `object` key.
 */
export function mapDriveItemsToQueryResults(
    items: DriveItem[],
    hasMoreItems: boolean = false,
    numItems?: number,
    succinct: boolean = true,
    includeAllowableActions: boolean = false
): { results: ObjectData[]; hasMoreItems: boolean; numItems?: number } {
    return {
        results: items.map(item => mapDriveItemToObjectData(item, succinct, includeAllowableActions)),
        hasMoreItems,
        numItems,
    };
}
