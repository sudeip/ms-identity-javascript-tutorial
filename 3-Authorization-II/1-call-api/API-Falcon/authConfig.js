const passportConfig = {
    credentials: {
        tenantID: "e5d00c4d-eacb-4b8c-bd6e-4e57a780ce5d",
        clientID: "07b9051c-9b0c-4370-b53b-89c430ad1c9e"
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
