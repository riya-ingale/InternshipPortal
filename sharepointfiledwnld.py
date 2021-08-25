from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File

url = "https://viteduin59337.sharepoint.com/sites/CMPN_5_ProfessionalCommunicationandEthics-II_MickyBarua/Student%20Work/Forms/AllItems.aspx?FolderCTID=0x012000DDE874B50A2D98468B363C56116B5810&id=%2Fsites%2FCMPN%5F5%5FProfessionalCommunicationandEthics%2DII%5FMickyBarua%2FStudent%20Work%2FWorking%20files%2FRiya%20Ingale%2FAssignment%201%5F%20Draft%20a%20resume%20and%20cover%20letter%2FRiya%27s%20Resume%20%281%29%2Epdf&parent=%2Fsites%2FCMPN%5F5%5FProfessionalCommunicationandEthics%2DII%5FMickyBarua%2FStudent%20Work%2FWorking%20files%2FRiya%20Ingale%2FAssignment%201%5F%20Draft%20a%20resume%20and%20cover%20letter"

ctx_auth = AuthenticationContext(url)
ctx_auth.acquire_token_for_user("riya.ingale@vit.edu.in", "password")
ctx = ClientContext(url, ctx_auth)
response = File.open_binary(ctx, "/Shared Documents/downloadedfile.pdf")
with open("./downloadedfile.pdf", "wb") as local_file:
    local_file.write(response.content)
