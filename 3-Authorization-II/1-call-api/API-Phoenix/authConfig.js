const passportConfig = {
    credentials: {
        tenantID: "e5d00c4d-eacb-4b8c-bd6e-4e57a780ce5d",
        clientID: "58d35cac-7505-4ab4-82e0-babaa572cb21"
    },
    metadata: {
        authority: "login.microsoftonline.com",
        discovery: ".well-known/openid-configuration",
        version: "v2.0"
    },
    settings: {
        validateIssuer: true,
        passReqToCallback: true,
        loggingLevel: "info",
        loggingNoPII: true,
    },
    protectedRoutes: {
        data: {
            endpoint: "/api/data",
            delegatedPermissions: {
                read: ["access_as_user"]
            }
        }
    }
}

module.exports = passportConfig;
