const API = "https://script.google.com/macros/s/AKfycbzqYuTEireOQPmAvygQNekdFyu_JuKkdKCju-ySMbOCFfbXXTVt03G8Mc8EEnWtPrJ7PA/exec";

const loginView = document.getElementById("login-view");
const teamView = document.getElementById("team-view");
const loginButton = document.getElementById("login-button");
const logoutButton = document.getElementById("logout-button");
const loginError = document.getElementById("login-error");

const teamIdInput = document.getElementById("team-id");
const adminCodeInput = document.getElementById("admin-code");

const playersDiv = document.getElementById("players");
const teamName = document.getElementById("team-name");
const teamLeague = document.getElementById("team-league");
const playerCount = document.getElementById("player-count");
const teamLogo = document.getElementById("team-logo");
const search = document.getElementById("player-search");

let currentPlayers = [];

loginButton.addEventListener("click", login);
logoutButton.addEventListener("click", logout);

search.addEventListener("input", () => {
  const value = search.value.toLowerCase();

  renderPlayers(
    currentPlayers.filter(player =>
      player.name.toLowerCase().includes(value)
    )
  );
});

async function login() {
  loginError.textContent = "";

  const teamId = teamIdInput.value.trim();
  const adminCode = adminCodeInput.value.trim();

  if (!teamId || !adminCode) {
    loginError.textContent = "Please enter your Team ID and Admin Code.";
    return;
  }

  loginButton.disabled = true;
  loginButton.textContent = "Signing in...";

  try {
    const url =
      API +
      "?type=login" +
      "&teamId=" + encodeURIComponent(teamId) +
      "&adminCode=" + encodeURIComponent(adminCode);

    const res = await fetch(url);
    const data = await res.json();

    if (!data.success) {
      loginError.textContent = data.error || "Login failed.";
      return;
    }

    localStorage.setItem("mfflSessionId", data.sessionId);

    await loadSession();

  } catch (err) {
    loginError.textContent = "Could not connect to the portal. Please try again.";
  } finally {
    loginButton.disabled = false;
    loginButton.textContent = "Sign In";
  }
}

async function loadSession() {
  const sessionId = localStorage.getItem("mfflSessionId");

  if (!sessionId) return;

  const url =
    API +
    "?type=session" +
    "&sessionId=" + encodeURIComponent(sessionId);

  const res = await fetch(url);
  const data = await res.json();

  if (!data.success) {
    localStorage.removeItem("mfflSessionId");
    loginView.classList.remove("hidden");
    teamView.classList.add("hidden");
    return;
  }

  showTeam(data);
}

function showTeam(data) {
  loginView.classList.add("hidden");
  teamView.classList.remove("hidden");

  teamName.textContent = data.team.name;
  teamLeague.textContent = data.team.league;
  playerCount.textContent = data.players.length + " Approved Players";

  if (data.team.logo) {
    teamLogo.style.backgroundImage = `url(${data.team.logo})`;
  } else {
    teamLogo.style.backgroundImage = "";
  }

  currentPlayers = data.players;
  search.value = "";
  renderPlayers(currentPlayers);
}

function logout() {
  localStorage.removeItem("mfflSessionId");
  currentPlayers = [];
  playersDiv.innerHTML = "";
  loginView.classList.remove("hidden");
  teamView.classList.add("hidden");
}

function renderPlayers(players) {
  playersDiv.innerHTML = "";

  players.forEach(player => {
    const card = document.createElement("div");
    card.className = "player-card";

    const bmi = player.showBMI
      ? `<div class="player-bmi"><strong>BMI:</strong> ${player.bmi}</div>`
      : "";

    const note = player.note
      ? `<div class="league-note"><strong>League Note</strong>${player.note}</div>`
      : "";

    card.innerHTML = `
      <div class="player-photo" style="background-image:url('${player.photo}')"></div>

      <div class="player-info">
        <h2>${player.name}</h2>
        <div class="player-id">${player.playerId}</div>
        <div class="player-route">${player.route}</div>
        ${bmi}
        ${note}
      </div>
    `;

    playersDiv.appendChild(card);
  });
}

loadSession();
