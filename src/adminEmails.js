export const ADMIN_EMAILS = [
  "tthomson@nationsfinest.org",
  "cflaherty@nationsfinest.org",
  "lwilliams@nationsfinest.org",
];

export function isAdminEmail(email) {
  return ADMIN_EMAILS.includes((email || "").trim().toLowerCase());
}
