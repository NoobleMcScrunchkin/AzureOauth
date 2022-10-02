import express from "express";
import dotenv from "dotenv";
import fetch from "node-fetch";
import fs from "fs";
import fs from "https";

dotenv.config();

var privateKey = fs.readFileSync("privatekey.pem");
var certificate = fs.readFileSync("certificate.pem");

const app = express();
const port = 2048;

app.get("/microsoft/callback", async (req, res) => {
	if (req.query.error && req.query.error_description) {
		console.error("Error: " + req.query.error + "\n" + req.query.error_description);
		res.send(false);
		return;
	}

	let token = await get_token(req.query.code);

	res.json(token);
});

app.get("/microsoft/refresh", async (req, res) => {
	if (!(req.query.refresh_token && req.query.refresh_token != "")) {
		console.error("Error: Invalid Refresh Token");
		res.send(false);
		return;
	}

	let token = await refresh_token(req.query.refresh_token);

	res.json(token);
});

https
	.createServer(
		{
			key: privateKey,
			cert: certificate,
		},
		app
	)
	.listen(port);

async function get_token(code) {
	let res = await fetch("https://login.live.com/oauth20_token.srf", {
		method: "POST",
		headers: {
			"Content-Type": "application/x-www-form-urlencoded",
		},
		body: "client_id=" + process.env.CLIENT_ID + "&redirect_uri=" + process.env.REDIRECT_URI + "&client_secret=" + process.env.CLIENT_SECRET + "&code=" + code + "&grant_type=authorization_code",
	});

	let json = await res.json();

	if (json.error) {
		console.error("Error: " + json.error + "\n" + json.error_description);
		return false;
	}

	return json;
}

async function refresh_token(refresh_token) {
	let res = await fetch("https://login.live.com/oauth20_token.srf", {
		method: "POST",
		headers: {
			"Content-Type": "application/x-www-form-urlencoded",
		},
		body: "client_id=" + process.env.CLIENT_ID + "&refresh_token=" + refresh_token + "&redirect_uri=" + process.env.REDIRECT_URI + "&client_secret=" + process.env.CLIENT_SECRET + "&grant_type=refresh_token",
	});

	let json = await res.json();

	if (json.error) {
		console.error("Error: " + json.error + "\n" + json.error_description);
		return false;
	}

	return json;
}
