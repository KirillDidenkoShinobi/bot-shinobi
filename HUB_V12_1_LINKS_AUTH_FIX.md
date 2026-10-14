# Shinobi Team Hub V12.1 — auth stability and invite links

- Moved blocking auth database calls off the shared asyncio event loop used by Telegram polling.
- Run the auth schema setup once per process instead of once per authenticated request.
- Added an administrator-only action in the Invitations and Growth page to create a personal invite link for the main chat. Links use join requests and are recorded under the creating Telegram ID.
- Added a red Create Link button and kept the existing link list available on the same page.

Roles allowed to create links: owner, deputy, admin. Telegram must have permission to create invite links in the configured main chat.
