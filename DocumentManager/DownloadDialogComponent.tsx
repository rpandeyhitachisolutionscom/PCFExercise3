/* eslint-disable @typescript-eslint/no-explicit-any */
import { Dialog } from '@fluentui/react';
import * as React from 'react';
import './main.css';

export const DownloadDialogComponent: React.FC<any> = ({ isDDialogOpen, onClose, allFiles, downloadFile }) => {

    React.useEffect(() => {
        console.log('rpshan');
    }, []);
    const download = (data: any) => {
        downloadFile(data);
    }
    return (
        <>
            <Dialog
                hidden={!isDDialogOpen}
                onDismiss={onClose}
                // modalProps={modalProps}
                styles={{
                    main: {
                        selectors: {
                            ['@media (min-width: 480px)']: {
                                minWidth: 450,
                                maxWidth: '1000px',
                                width: '1000px',
                                padding: '30px'
                            }
                        }
                    }
                }}
            >
                <div>
                    <h3 className='alc'> For Download File Click On That</h3>
                    <div className='list_file dflexd'>

                        {allFiles.map((file: any) => (
                            <button className='dbutton' onClick={() => download(file)}>{file}</button>
                        ))}
                    </div>
                </div>
            </Dialog>
        </>
    )


}