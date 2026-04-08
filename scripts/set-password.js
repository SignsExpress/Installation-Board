const { ensureUsersFile, getUsersFile, setUserPassword } = require("../server/auth-store");

async function main() {
  const [, , displayName, password] = process.argv;
  if (!displayName || !password) {
    console.error('Usage: node scripts/set-password.js "Display Name" "NewPassword"');
    process.exit(1);
  }

  ensureUsersFile();
  const user = await setUserPassword(displayName, password);
  console.log(`Password set for ${user.displayName} (${user.role}).`);
  console.log(`Users file: ${getUsersFile()}`);
}

main().catch((error) => {
  console.error(error.message || error);
  process.exit(1);
});
