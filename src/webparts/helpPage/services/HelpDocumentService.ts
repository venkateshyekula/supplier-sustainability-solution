import {
    WebPartContext
} from '@microsoft/sp-webpart-base';

import {
    SPFI
} from '@pnp/sp';

import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

import {
    getSP
} from '../../../pnpConfig';

import {
    IHelpDocument
} from '../models/IHelpDocument';

import {
    IHelpDocumentLinks
} from '../models/IHelpDocumentLinks';

interface IHelpDocumentListItem {
    Id: number;
    Title?: string;
    FileLeafRef?: string;
    FileRef?: string;
    Modified?: string;
    ResourceKey?: string;
    IsActive?: boolean;
    FSObjType?: number;
}

export class HelpDocumentService {
    private readonly sp:
        SPFI;

    public constructor(
        context: WebPartContext
    ) {
        this.sp =
            getSP(context);
    }

    public async getDocumentLinks(
        libraryTitle: string
    ): Promise<IHelpDocumentLinks> {
        const documents:
            readonly IHelpDocument[] =
            await this.getDocuments(
                libraryTitle
            );

        return {
            esgGuidelinesUrl:
                this.findUrl(
                    documents,
                    'EsgGuidelines'
                ),

            userGuideUrl:
                this.findUrl(
                    documents,
                    'UserGuide'
                ),

            policyComplianceUrl:
                this.findUrl(
                    documents,
                    'PolicyCompliance'
                ),

            supplierGuideUrl:
                this.findUrl(
                    documents,
                    'SupplierGuide'
                ),

            procurementGuideUrl:
                this.findUrl(
                    documents,
                    'ProcurementGuide'
                ),
            faqPageUrl:
                this.findUrl(
                    documents,
                    'FAQ'
                ),

            supplierPortalUrl:
                this.findUrl(
                    documents,
                    'SupplierPortalGuide'
                )
        };
    }

    private async getDocuments(
        libraryTitle: string
    ): Promise<IHelpDocument[]> {
        const normalizedLibraryTitle:
            string =
            libraryTitle
                ? libraryTitle.trim()
                : '';

        if (!normalizedLibraryTitle) {
            throw new Error(
                'The Help document library title is not configured.'
            );
        }

        const items:
            IHelpDocumentListItem[] =
            await this.sp.web.lists
                .getByTitle(
                    normalizedLibraryTitle
                )
                .items
                .select(
                    'Id',
                    'Title',
                    'FileLeafRef',
                    'FileRef',
                    'Modified',
                    'ResourceKey',
                    'IsActive',
                    'FSObjType'
                )
                .filter(
                    'FSObjType eq 0'
                )
                .orderBy(
                    'Modified',
                    false
                )
                .top(
                    100
                )() as IHelpDocumentListItem[];

        return items
            .filter(
                (
                    item:
                        IHelpDocumentListItem
                ): boolean => {
                    return Boolean(
                        item.ResourceKey &&
                        item.ResourceKey.trim() &&
                        item.FileRef &&
                        item.FileRef.trim()
                    );
                }
            )
            .map(
                (
                    item:
                        IHelpDocumentListItem
                ): IHelpDocument => {
                    const fileName: string =
                        item.FileLeafRef
                            ? item.FileLeafRef.trim()
                            : '';

                    return {
                        id:
                            item.Id,

                        resourceKey:
                            item.ResourceKey
                                ? item.ResourceKey.trim()
                                : '',

                        title:
                            item.Title &&
                                item.Title.trim()
                                ? item.Title.trim()
                                : this.removeFileExtension(
                                    fileName
                                ),

                        fileName,

                        serverRelativeUrl:
                            item.FileRef
                                ? item.FileRef.trim()
                                : '',

                        modified:
                            item.Modified
                    };
                }
            );
    }

    private findUrl(
        documents:
            readonly IHelpDocument[],
        resourceKey:
            string
    ): string {
        const normalizedResourceKey:
            string =
            resourceKey
                .trim()
                .toLowerCase();

        const matchingDocument:
            IHelpDocument | undefined =
            documents.find(
                (
                    document:
                        IHelpDocument
                ): boolean => {
                    return (
                        document.resourceKey
                            .trim()
                            .toLowerCase() ===
                        normalizedResourceKey
                    );
                }
            );

        return matchingDocument
            ? matchingDocument.serverRelativeUrl
            : '';
    }

    private removeFileExtension(
        fileName: string
    ): string {
        const lastDotIndex: number =
            fileName.lastIndexOf('.');

        if (lastDotIndex <= 0) {
            return fileName;
        }

        return fileName.substring(
            0,
            lastDotIndex
        );
    }
}