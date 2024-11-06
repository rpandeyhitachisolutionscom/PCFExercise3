/* eslint-disable @typescript-eslint/no-explicit-any */
export const getData = async (clientOrgURL: string, requestApiQuery?: string):Promise<any> => {
    const clientUrl = clientOrgURL;
    //const queryString = query ? `?${query}` : "";
    const requestUrl = `${clientUrl}/api/data/v9.1/${requestApiQuery}`;
    try {
        const response = await fetch(requestUrl, {
            method: "GET",
            headers: {
                "OData-MaxVersion": "4.0",
                "OData-Version": "4.0",
                "Accept": "application/json",
                "Content-Type": "application/json; charset=utf-8",
                "Prefer": "odata.include-annotations=OData.Community.Display.V1.FormattedValue"
            }
        });

        if (!response.ok) {
            throw new Error(`Error fetching data: ${response.statusText}`);
        }

        const data = await response.json();
       
        console.log(data);
        return data;
    } catch (error) {
        console.error("Error in getData:", error);
        throw error;
    }
}