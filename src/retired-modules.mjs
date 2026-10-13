// Retire page entry points only; shared APIs and saved operational data remain intact.
export function isRetiredModulePath(pathname) {
 return /^\/(?:client\/)?(?:morning-meeting|filtering|materials|mustang)(?:\/|$)/i.test(pathname);
}
