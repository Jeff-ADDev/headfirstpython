

#    "self": "https://revlocaldev.atlassian.net/rest/api/2/user?accountId=5ae8c5b7347761310e255f23",
#    "accountId": "5ae8c5b7347761310e255f23",
#    "accountType": "atlassian",
#    "emailAddress": "msullivan@revlocal.com",
#    "avatarUrls": {
#        "48x48": "https://avatar-management--avatars.us-west-2.prod.public.atl-paas.net/5ae8c5b7347761310e255f23/1a6fa6f2-34bf-4e8b-8cd8-96cc905941af/48",
#        "24x24": "https://avatar-management--avatars.us-west-2.prod.public.atl-paas.net/5ae8c5b7347761310e255f23/1a6fa6f2-34bf-4e8b-8cd8-96cc905941af/24",
#        "16x16": "https://avatar-management--avatars.us-west-2.prod.public.atl-paas.net/5ae8c5b7347761310e255f23/1a6fa6f2-34bf-4e8b-8cd8-96cc905941af/16",
#        "32x32": "https://avatar-management--avatars.us-west-2.prod.public.atl-paas.net/5ae8c5b7347761310e255f23/1a6fa6f2-34bf-4e8b-8cd8-96cc905941af/32"
#    },
#    "displayName": "Matt Sullivan",
#    "active": true,
#    "locale": "en_US"

class User:
    def __init__(self, id, email, name, active):
        self.id = id
        self.email = email
        self.name = name
        self.active = active
        