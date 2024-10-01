// Create HTTP server.
const restify = require("restify");
const send = require("send");
const { commandApp } = require("./internal/initialize");
const { TeamsBot } = require("./teamsBot");

// This template uses `restify` to serve HTTP responses.
// Create a restify server.
const server = restify.createServer();
server.use(restify.plugins.bodyParser());

server.get(
  "/static/*",
  restify.plugins.serveStatic({
    directory: __dirname,
  })
);

server.listen(process.env.port || process.env.PORT || 3978, () => {
  console.log(`\nApp Started, ${server.name} listening to ${server.url}`);
});

// Setup the static tab
server.get("/tab", (req, res, next) => {
  send(req, __dirname + "/views/hello.html").pipe(res);
});

// Register an API endpoint with `restify`. Teams sends messages to your application
// through this endpoint.
//
// The Teams Toolkit bot registration configures the bot with `/api/messages` as the
// Bot Framework endpoint. If you customize this route, update the Bot registration
// in `templates/azure/provision/botservice.bicep`.
const teamsBot = new TeamsBot();
server.post("/api/messages", async (req, res) => {
  await commandApp.requestHandler(req, res, async (context) => {
    await teamsBot.run(context);
  });
});
