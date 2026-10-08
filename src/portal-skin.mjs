export const SKIN_KEY = 'sx-portal-skin';
export function resolveSkin(search, stored) {
  const requested = new URLSearchParams(search).get('skin');
  if (requested === 'new' || requested === 'classic') return requested;
  return stored === 'new' ? 'new' : 'classic';
}
export function skinLink(href, skin) {
  const url = new URL(href);
  url.searchParams.set('skin', skin === 'new' ? 'new' : 'classic');
  return url.href;
}
