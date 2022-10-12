var axios = require('axios');
var qs = require('qs');
const dAssetObjectTypes = ["Files", "Folders", "Glossary Terms", "Reports", "Stored procedures", "Tables", "Dashboards", "Data pipelines"]

async function getBearerToken()
{
    try{
        var data = qs.stringify({
            'grant_type': 'client_credentials',
            'client_id': 'a5231317-4c9c-4820-af8f-6bc51d343137',
            'client_secret': 'wPA8Q~RNSKJHPWQXoaYWUzAiQdhiKWAWQCLHMb9Y',
            'resource\t': 'https://purview.azure.net' 
            });
            var config = {
            method: 'post',
            url: 'https://login.microsoftonline.com/bc1fe3f1-755f-402f-a2b1-9f20f20b860c/oauth2/token',
            headers: { 
                'Content-Type': 'application/x-www-form-urlencoded', 
                'Cookie': 'fpc=AnJJs97yToBPlDFeBRokozEQDLcQAQAAAJ3M1doOAAAA; stsservicecookie=estsfd; x-ms-gateway-slice=estsfd'
            },
            data : data
            };
            let result = await axios(config);
            if(result.status == 200)
            {
                console.log(result.status+" Request for Bearertoken was successful");
            }
            return result.data.access_token;
    }
    catch (err)
    {
        console.log(err)
    }
}

async function searchForGuid(bToken, searchTerm, type)
{
    try{
        var data = JSON.stringify({
            'keywords' : searchTerm
        });

        var config = {
            method : 'post',
            url : 'https://pvlab-pv.purview.azure.com/catalog/api/search/suggest?api-version=2022-03-01-preview',
            headers : {
                'Authorization' : 'Bearer '+bToken,
                'Content-Type': 'application/json'
            },
            data : data
        };
        let result = await axios(config);
        if(result.status == 200)
        {
            console.log(result.status+" Request for "+type+" ID was successful");
        }
        for(var i=0; i < result.data.value.length; i++)
        {
            switch(type)
            {
                case("Glossary_Term"):{
                    if(result.data.value[i].name == searchTerm && result.data.value[i].objectType=="Glossary terms")
                        {
                            return result.data.value[i].id;
                        }
                }
                case("Data_Asset"):{
                    if(result.data.value[i].name == searchTerm && dAssetObjectTypes.includes(result.data.value[i].objectType))
                    {
                        return result.data.value[i].id;
                    }
                }
            }
        }
    }
    catch (err)
    {
        console.log(err);
    }
}
async function assignTermToAsset(bToken, gTermId, dAssetId)
{
    try{
        var data = JSON.stringify([
            {
            "guid": dAssetId
            }
        ]);
        
        var config = {
            method: 'post',
            url: 'https://pvlab-pv.purview.azure.com/catalog/api/atlas/v2/glossary/terms/'+gTermId+'/assignedEntities',
            headers: { 
            'Authorization': 'Bearer '+bToken, 
            'Content-Type': 'application/json'
            },
            data : data
        };
        
        let result = await axios(config);
        if(result.status == 204)
        {
            console.log(result.status+" The request for the assignment was done")
        }
    }
    catch (err) {
        console.log(err);
    }
}
module.exports = async function (context, req) {
    try {   
        
        // create a response message if the parameters are not set
        let responseMessage = "";
        if(typeof req.query.gterm == "undefined")
        {
            responseMessage += "The glossary term was not given in the parameters \n";
        }
        if(typeof req.query.dasset == "undefined")
        {
            responseMessage += "The data asset name was not given in the parameters \n";
        }

        const GlossaryTerm = req.query.gterm ? req.query.gterm : null;
        const DataAssetName = req.query.dasset ? req.query.dasset : null; 

        let bToken;
        let gTermId;
        let dAssetId;
        if(GlossaryTerm != null && DataAssetName != null){
            bToken = await getBearerToken();
            gTermId = await searchForGuid(bToken, GlossaryTerm, "Glossary_Term");
            dAssetId = await searchForGuid(bToken, DataAssetName, "Data_Asset");
        }

        // validate if the id's are set and also different
        if(gTermId && dAssetId && gTermId != dAssetId){
            await assignTermToAsset(bToken, gTermId, dAssetId);
            responseMessage += "The term "+ GlossaryTerm + " was assigned to the data asset "+ DataAssetName;
            console.log("The Term was successfully assigned!");
        }
        else if (typeof gTermId == 'undefined') {
            console.log("Requestet glossary term was not found..");
        }
        else{
            console.log("Requestet data asset was not found..");
        }
       
        context.res = {
            body: responseMessage
        };
    }catch (err){
        console.log(err);
    }
}
