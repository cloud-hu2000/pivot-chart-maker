export function FooterNetworkLinks({ links }) {
  if (!links.length) return null;
  return (
    <nav aria-label="More websites" data-footer-network="true" style={{ display: 'flex', flexWrap: 'wrap', gap: '12px', padding: '12px 0' }}>
      {links.map((link) => (
        <a key={link.url} href={link.url} target="_blank" rel="noopener noreferrer">{link.label}</a>
      ))}
    </nav>
  );
}
