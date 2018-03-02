import onedrivesdk
from onedrivesdk.helpers import GetAuthCodeServer

redirect_uri = 'http://localhost:500/'
client_secret = 'meQZOPPG70_$=pcesgV283}'
scopes=['wl.signin', 'wl.offline_access', 'onedrive.readwrite']

client = onedrivesdk.get_default_client(
    client_id='7b57ad44-02fe-4aca-b588-4033e2f66885', scopes=scopes)

auth_url = client.auth_provider.get_auth_url(redirect_uri)

#this will block until we have the code
code = GetAuthCodeServer.get_auth_code(auth_url, redirect_uri)

client.auth_provider.authenticate(code, redirect_uri, client_secret)
