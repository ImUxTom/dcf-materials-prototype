# allocate-cases

## Local setup

1. Clone the repo and switch to the branch you want to work on.
2. Open the project in your terminal and run these commands one at a time, in this order:

```bash
npm install
npx prisma generate
npm run generate-data
npm run dev
```

3. Open the app in your browser at:
- http://localhost:3000
- http://localhost:3000/manage-prototype

If you run into issues, make sure Node.js and npm are installed first, and if the prototype kit says it cannot find the command, run `npm install` again.

