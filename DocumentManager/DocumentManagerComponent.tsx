/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import { IInputs } from './generated/ManifestTypes';
import { useCallback, useEffect, useState } from 'react';
import { Dialog, Icon } from '@fluentui/react';
import { BlobServiceClient } from '@azure/storage-blob';
import { Buffer } from 'buffer';
import './main.css';
import { v4 as uuidv4 } from 'uuid';
import { DownloadDialogComponent } from './DownloadDialogComponent';
import { Spinner, SpinnerSize } from '@fluentui/react';
import {getData} from './DynamicService';
export interface DocumentManagerComponentProps {
    label: string;
    onChanges: (newValue: string | undefined) => void;
    context: ComponentFramework.Context<IInputs>;
    sasToken: string;
    containerName: string;
}
interface FileUrl {
    file: File,
    preview: string
}
//environment variable in crm , 

export const DocumentManagerComponent = React.memo((props: DocumentManagerComponentProps) => {
    const [isDialogOpen, setIsDialogOpen] = useState(false);
    const openDialog = (): any => setIsDialogOpen(true);
    const closeDialog = (): any => setIsDialogOpen(false);
    const [isDDialogOpen, setIsDDialogOpen] = useState(false);
    const openDDialog = (): any => setIsDDialogOpen(true);
    const closeDDialog = (): any => setIsDDialogOpen(false);
    const [fileName, setFileName] = useState<any>('');
    const [previewUrl, setPreviewUrl] = useState<string | null>(null);
    const [file, setFile] = useState<File | null>(null);
    const [files, setFiles] = useState<File[]>([]);
    const [allFiles, setAllFiles] = useState<any>([]);
    // const containerName = 'filecontainer';
    const [fileAndUrl, setFileAndUrl] = useState<FileUrl[]>([]);
    //const sasToken = 'sv=2022-11-02&ss=bfqt&srt=sco&sp=rwdlacupiytfx&se=2024-10-31T15:31:35Z&st=2024-10-23T07:31:35Z&spr=https&sig=i7bSNQ5gLNwXpaNqoU1TSHGkb%2BNulZL2upH%2FQmmzkjc%3D';
    const [containerName, setContainerName] = useState('');
    const [sasToken, setSasToken] = useState<string>('');

    // let blobServiceClient: any;
    // let containerClient: any;
    const [blobServiceClient, setBlobServiceClient] = useState<BlobServiceClient | null>(null);
    const [containerClient, setContainerClient] = useState<any>(null);
    // const blobServiceClient = new BlobServiceClient(`https://roshan.blob.core.windows.net/?${sasToken}`);
    // const containerClient = blobServiceClient.getContainerClient(containerName);
    const [isLoading, setIsLoading] = useState(false);
    const [nguid, setNguid] = useState<any>(null);
    useEffect(() => {
        window.Buffer = Buffer;
       // getData('https://org01fafc2a.api.crm.dynamics.com', 'environmentvariabledefinitions?$select=displayname,defaultvalue');
        if (props.label) {
            // setFileName(props.label);
            // const all = props.label.split(',');
            // setAllFiles(all);
            setNguid(props.label);
             fetchFilesByGuid(props.label);
        }
    }, [])


    // const getData = async (clientOrgURL: string, requestApiQuery?: string) => {
    //     const clientUrl = clientOrgURL;
    //     //const queryString = query ? `?${query}` : "";
    //     const requestUrl = `${clientUrl}/api/data/v9.1/${requestApiQuery}`;
    //     try {
    //         const response = await fetch(requestUrl, {
    //             method: "GET",
    //             headers: {
    //                 "OData-MaxVersion": "4.0",
    //                 "OData-Version": "4.0",
    //                 "Accept": "application/json",
    //                 "Content-Type": "application/json; charset=utf-8",
    //                 "Prefer": "odata.include-annotations=OData.Community.Display.V1.FormattedValue"
    //             }
    //         });

    //         if (!response.ok) {
    //             throw new Error(`Error fetching data: ${response.statusText}`);
    //         }

    //         const data = await response.json();
    //         const sas = data.value.find((d: any) => d.displayname === 'BlobSasToken')?.defaultvalue;
    //         const con = data.value.find((d: any) => d.displayname === 'BlobContainer')?.defaultvalue;

    //         if (sas && con) {
    //             setSasToken(sas);
    //             setContainerName(con);
    //             localStorage.setItem('sasToken', sas);
    //             localStorage.setItem('containerName', con);
    //             initializeBlobClients(sas, con);
    //             fetchFilesByGuid(props.label);
    //         }
    //         console.log(data);
    //         return data;
    //     } catch (error) {
    //         console.error("Error in getData:", error);
    //         throw error;
    //     }
    // }

    const initializeBlobClients = (sas: string, container: string): any => {
        const blobServiceClientInstance = new BlobServiceClient(`https://roshan.blob.core.windows.net/?${props.sasToken}`);
        const containerClientInstance = blobServiceClientInstance.getContainerClient(props.containerName||'');
        return containerClientInstance;
    };

    const fetchFilesByGuid = async (guid: any) => {
        const containerClient = initializeBlobClients('', '');
        const blobs = containerClient.listBlobsFlat({ prefix: `${guid}/` });
        const files = [];
        for await (const blob of blobs) {
            files.push(blob.name); // Store the blob names or other metadata
        }
        setAllFiles(files); // This will return an array of file names uploaded under the specific GUID
    };

    const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
        if (event.target.files) {
            const selectedFile = event.target.files[0];
            setFile(event.target.files[0]);
            const objectUrl = URL.createObjectURL(selectedFile);
            //setPreviewUrl(objectUrl);
        }
    };

    const handleFilesChange = (event: React.ChangeEvent<HTMLInputElement>) => {
        if (event.target.files) {
            const selectedFiles = Array.from(event.target.files);
            const allowedExtensions = ['jpg', 'jpeg', 'png', 'pdf', 'doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'txt'];

            // Filter out files that do not match allowed extensions
            const filteredFiles = selectedFiles.filter(file => {
                const extension = file.name.split('.').pop()?.toLowerCase();
                return extension && allowedExtensions.includes(extension);
            });

            if (filteredFiles.length === 0) {
                alert('No valid files selected. Please choose JPG, PNG, Word, PDF, Excel, PPTX, or TXT files.');
                return;
            }

            setFiles(selectedFiles);
            const urls: any = [];
            selectedFiles.forEach((file: any) => {
                const objectUrl = URL.createObjectURL(file);
                urls.push({ file: file, preview: objectUrl });
            })
            //setPreviewUrl(urls);
            setFileAndUrl(urls);
        }
    };

    const uploadFiles = useCallback(async () => {
        if (files.length > 0) {
            setIsLoading(true);
            const uniqueGuid = (nguid == null) ? (uuidv4()) : (nguid);
            const EntityName = props.context.parameters.ContainerName.raw;
            let tempFileName = '';
            try {
                await Promise.all(files.map(async (file) => {
                    // const uniqueFileName = `${file.name.split('.').slice(0, -1).join('.')}_${Date.now()}.${file.name.split('.').pop()}`;
                   const folderPath = `${EntityName}/${uniqueGuid}/`
                    const uniqueFileName = `${folderPath}${file.name}`;
                    const containerClient = initializeBlobClients('','');
                    const blockBlobClient = containerClient.getBlockBlobClient(uniqueFileName);
                    const options = {
                        blobHTTPHeaders: { blobContentType: file.type },    
                    };
                    await blockBlobClient.uploadData(file, options);
                    console.log(`${file.name} uploaded successfully`);
                    tempFileName = tempFileName + uniqueFileName + ',';
                    const files = allFiles;
                    files.push(uniqueFileName);
                    setAllFiles(files);
                }));
                setFiles([]); // Clear the file input after upload
                tempFileName = tempFileName.slice(0, -1);
                setFileName(tempFileName);
                props.onChanges(uniqueGuid);
                setIsLoading(false);
                alert("Document Uploade Successfully");
                closeDialog();
            } catch (error) {
                console.log(error);
            }
        }
    }, [files, containerClient]);

    // const uploadFile = useCallback(async () => {
    //     if (file) {
    //         try {
    //             const uniqueFileName = `${file.name}_${Date.now()}`;
    //             // const blockBlobClient = containerClient.getBlockBlobClient(file.name);
    //             const blockBlobClient = containerClient.getBlockBlobClient(uniqueFileName);
    //             const options = {
    //                 blobHTTPHeaders: { blobContentType: file.type }, // Set the content type of the blob
    //             };
    //             // Use the upload method instead
    //             await blockBlobClient.uploadData(file, options);
    //             console.log('File uploaded successfully');
    //             setFile(null);
    //             // setFileName(file.name);
    //             // props.onChanges(file.name);
    //             setFileName(uniqueFileName);
    //             props.onChanges(uniqueFileName);
    //         } catch (error) {
    //             console.log(error);
    //         }
    //     }
    // }, [file, containerClient]);

    const downloadFile = async (fileItem: any) => {
        const blockBlobClient = containerClient.getBlockBlobClient(fileItem);
        const url = blockBlobClient.url;
        // window.open(url, '_blank');
        try {
            const response = await fetch(url);
            if (!response.ok) {
                throw new Error('Network response was not ok');
            }
            const blob = await response.blob();
            // Create a link element to trigger the download
            const link = document.createElement('a');
            link.href = window.URL.createObjectURL(blob);
            link.download = fileItem.split('/').pop(); // Set the desired file name for download
            // Append to the body (required for Firefox)
            document.body.appendChild(link);
            // Trigger the download
            link.click();
            // Clean up and remove the link
            document.body.removeChild(link);
        } catch (error) {
            console.error('Download failed:', error);
        }
    }

    const handleDeleteFile = (file: any) => {
        const confirmDelete = window.confirm("Are you sure you want to delete this file?");
        if (confirmDelete) {
            deleteFile(file);
        }
    };

    const deleteFile = async (uniqueFileName: any) => {
        const containerClient = initializeBlobClients('', '');
        const blockBlobClient = containerClient.getBlockBlobClient(uniqueFileName);
        try {
            await blockBlobClient.delete();
            let files = [];
            files = allFiles.filter((d: any) => d != uniqueFileName);
            setAllFiles(files);
            console.log(`${uniqueFileName} deleted successfully`);
        } catch (error) {
            console.error(`Error deleting ${uniqueFileName}:`, error);
        }
    };

    const previewFile = useCallback(() => {
        const blockBlobClient = containerClient.getBlockBlobClient(fileName);
        const url = blockBlobClient.url;
        //setPreviewUrl(url);
    }, [containerClient]);

    const pretext = (url: any) => {
        window.open(url || '', '_blank');
    }

    return (
        <>
            {/* //<input type='text' placeholder='Upload Resume File' onClick={openDialog} value={fileName} /> */}
            <div className='dflexc'>
                <button onClick={openDialog} className='dbutton'>Upload</button>
                <div className='dflexd mt-18'>
                    {allFiles?.map((file: any) => (
                        <div className='dflexi'>
                            <p className='textp' onClick={() => downloadFile(file)}>{file.split('/').pop()}</p>
                            <Icon
                                iconName="Delete"
                                className='icnf cnb '
                                onClick={() => handleDeleteFile(file)}
                            />
                        </div>
                    ))}
                </div>
            </div>
            <Dialog
                hidden={!isDialogOpen}
                onDismiss={closeDialog}
                // modalProps={modalProps}
                styles={{
                    main: {
                        selectors: {
                            ['@media (min-width: 480px)']: {
                                minWidth: 450,
                                maxWidth: '1000px',
                                width: '1000px',
                                padding: '30px',
                                height: '400px'
                            }
                        }
                    }
                }}
            >
                {isLoading ? <Spinner size={SpinnerSize.large} label="Loading..." /> :
                    <div className='main-container'>
                        <div className='top dflex' >
                            <div className='dflex gp'>
                                <div className='download dflex' style={{ display: fileName == '' ? 'none' : 'block' }}>
                                    {/* <div className='dflex cp'>
                                        <Icon
                                            iconName="Download"
                                            onClick={() => openDDialog()}
                                            className="icn"
                                        />
                                        <p onClick={() => openDDialog()} className='download_text' >download file</p>
                                    </div> */}
                                </div>
                                {/* <div className='prview dflex'>
                                <button onClick={() => previewFile()} disabled={!file}>Preview</button>
                            </div> */}
                            </div>
                            <Icon iconName="Cancel" onClick={closeDialog} className='icn cp fw-600' style={{ cursor: 'pointer', fontSize: '24px', color: 'red' }} />
                        </div>

                        <DownloadDialogComponent isDDialogOpen={isDDialogOpen} onClose={closeDDialog} allFiles={allFiles} downloadFile={downloadFile} />

                        <div className='upload'>
                            {/* <input type="file" onChange={handleFileChange} />
                        <button onClick={uploadFile} disabled={!file}>Upload</button> */}
                            <input type="file" onChange={handleFilesChange} multiple />
                            <button onClick={uploadFiles} disabled={!files} className='dbutton'>Upload</button>
                        </div>
                        {/* <div className='preview' style={{display:previewUrl !=null?'block':'none'}}>
                        <h3>Preview:</h3>
                        {file?.type.startsWith('image/') ? (
                            <img src={previewUrl || ''} alt="Preview" style={{ width: '100%', maxHeight: '400px' }} />
                        ) : file?.type === 'application/pdf' ? (
                            <iframe src={previewUrl || ''} style={{ width: '100%', height: '400px' }} />
                        ) : file?.type === 'text/plain' ? (
                            <pre onClick={pretext} className='linktext'>Click Here For Preview</pre>
                        ) :file?.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ? (
                            <iframe src={`https://view.officeapps.live.com/op/view.aspx?src=${encodeURIComponent(previewUrl||'')}`} style={{ width: '100%', height: '400px' }} />
                        ) : file?.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ? (
                            <iframe src={`https://view.officeapps.live.com/op/view.aspx?src=${encodeURIComponent(previewUrl||'')}`} style={{ width: '100%', height: '400px' }} />
                        ) : (
                            <p>Preview not available for this file type.</p>
                        )}
                    </div> */}

                        <div className='preview' style={{ display: fileAndUrl.length !== 0 ? 'block' : 'none' }}>
                            <h3>Preview:</h3>
                            {fileAndUrl.map((file, index) => (
                                <div key={index}>
                                    {file.file?.type.startsWith('image/') ? (
                                        <img src={file.preview || ''} alt="Preview" style={{ width: '100%', maxHeight: '400px' }} />
                                    ) : file.file?.type === 'application/pdf' ? (
                                        <iframe src={file.preview || ''} style={{ width: '100%', height: '400px' }} />
                                    ) : file.file?.type === 'text/plain' ? (
                                        <pre onClick={() => pretext(file.preview)} className='linktext'>Click Here For Preview</pre>
                                    ) : file.file?.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
                                        file.file?.type === 'application/msword' ? (
                                        <pre onClick={() => pretext(file.preview)} className='linktext'>Click Here For Preview</pre>
                                    ) : file.file?.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
                                        file.file?.type === 'application/vnd.ms-excel' ? (
                                        <pre onClick={() => pretext(file.preview)} className='linktext'>Click Here For Preview</pre>
                                    ) : (
                                        <p>Preview not available for this file type.</p>
                                    )}
                                </div>

                            ))}
                        </div>
                    </div>
                }
            </Dialog>
        </>
    );
});
