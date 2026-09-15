const { hasRequiredDelegatedPermissions } = require('../auth/permissionUtils');

const authConfig = require('../authConfig');

/**
 * Returns a small payload identifying this API (Falcon), so the BlueFlames
 * demo can show that the access token's 'aud' claim matches this app while
 * the caller's identity (from the token) stays the same across apps.
 */
exports.getData = (req, res, next) => {
    if (hasRequiredDelegatedPermissions(req.authInfo, authConfig.protectedRoutes.data.delegatedPermissions.read)) {
        try {
            res.status(200).json({
                app: 'Falcon',
                message: 'Hello from Falcon API',
                user: req.authInfo.preferred_username || req.authInfo.upn,
                time: new Date().toISOString(),
            });
        } catch (error) {
            next(error);
        }
    } else {
        next(new Error('User does not have the required permissions'));
    }
}
