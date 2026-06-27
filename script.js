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
let ownPlayers = [];
let oppositionPlayers = [];
let dashboardData = null;
let activeSquad = "own";

let squadSwitcher = document.getElementById("squad-switcher");

if (!squadSwitcher) {
  squadSwitcher = document.createElement("div");
  squadSwitcher.id = "squad-switcher";
  squadSwitcher.className = "squad-switcher hidden";

  search.parentNode.insertBefore(squadSwitcher, search);
}

let leagueView = document.getElementById("league-view");

if (!leagueView) {
  leagueView = document.createElement("section");
  leagueView.id = "league-view";
  leagueView.className = "hidden";
  leagueView.innerHTML = `
    <button id="league-logout-button" class="back-button">← Log out</button>
    <header class="team-header">
      <h1 id="league-title">MFFL Club Directory</h1>
      <p id="league-subtitle">Select a club to view approved players</p>
    </header>
    <div id="league-groups"></div>
  `;
  document.querySelector(".app").appendChild(leagueView);
}

const leagueGroups = document.getElementById("league-groups");
const leagueTitle = document.getElementById("league-title");
const leagueSubtitle = document.getElementById("league-subtitle");
const leagueLogoutButton = document.getElementById("league-logout-button");

loginButton.addEventListener("click", login);
leagueLogoutButton.addEventListener("click", logout);

teamIdInput.addEventListener("keydown", event => {
  if (event.key === "Enter") login();
});

adminCodeInput.addEventListener("keydown", event => {
  if (event.key === "Enter") login();
});

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
    loginError.textContent = "Please enter your ID and code.";
    return;
  }

  loginButton.disabled = true;
  loginButton.innerHTML = '<span class="spinner"></span> Signing in...';

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
    loginButton.innerHTML = "Sign In";
  }
}

async function loadSession() {
  const sessionId = localStorage.getItem("mfflSessionId");

  if (!sessionId) return;

  try {
    const url =
      API +
      "?type=session" +
      "&sessionId=" + encodeURIComponent(sessionId);

    const res = await fetch(url);
    const data = await res.json();

    if (!data.success) {
      console.warn("Session failed:", data.error);
      localStorage.removeItem("mfflSessionId");
      showLogin();
      return;
    }

    dashboardData = data;

    if (data.mode === "league") {
      showLeagueDashboard(data);
    } else {
      showTeam(data);
    }

  } catch (err) {
    localStorage.removeItem("mfflSessionId");
    showLogin();
  }
}

function showLogin() {
  loginView.classList.remove("hidden");
  teamView.classList.add("hidden");
  leagueView.classList.add("hidden");
}

function showLeagueDashboard(data) {
  loginView.classList.add("hidden");
  teamView.classList.add("hidden");
  leagueView.classList.remove("hidden");

  leagueTitle.textContent = "MFFL Club Directory";
  leagueSubtitle.textContent = "Select a club to view approved players";

  leagueGroups.innerHTML = "";

  const allTeams = data.groups.flatMap(group => group.teams);

  const grid = document.createElement("div");
  grid.className = "team-grid";

  allTeams.forEach(team => {
    const card = document.createElement("button");
    card.className = "team-card";

    card.innerHTML = `
      <img class="club-logo" src="${team.logo}" alt="${team.name}">
      <h2>${team.name}</h2>
      <p>${team.league}</p>
    `;

    card.addEventListener("click", () => loadTeamFromDashboard(team.slug));

    grid.appendChild(card);
  });

  leagueGroups.appendChild(grid);
}

async function loadTeamFromDashboard(slug) {
  const sessionId = localStorage.getItem("mfflSessionId");

  if (!sessionId) {
    showLogin();
    return;
  }

  try {
    const url =
      API +
      "?type=team" +
      "&sessionId=" + encodeURIComponent(sessionId) +
      "&slug=" + encodeURIComponent(slug);

    const res = await fetch(url);
    const data = await res.json();

    if (!data.success) {
      alert(data.error || "Could not load team.");
      return;
    }

    showTeam(data, true);

  } catch (err) {
    alert("Could not connect to the portal.");
  }
}

function showTeam(data, fromLeagueDashboard = false) {
  loginView.classList.add("hidden");
  leagueView.classList.add("hidden");
  teamView.classList.remove("hidden");

  teamName.textContent = data.team.name;
  teamLeague.textContent = data.team.league;

  ownPlayers = data.players || [];
  oppositionPlayers = data.opposition ? data.opposition.players || [] : [];
  activeSquad = "own";

  playerCount.textContent = ownPlayers.length + " Approved Players";

  if (data.team.logo) {
    teamLogo.style.backgroundImage = `url(${data.team.logo})`;
  } else {
    teamLogo.style.backgroundImage = "";
  }

  buildSquadSwitcher(data);

  currentPlayers = ownPlayers;
  search.value = "";
  renderPlayers(currentPlayers);

  logoutButton.textContent = fromLeagueDashboard ? "← Back to clubs" : "← Log out";

  logoutButton.onclick = async () => {
    if (fromLeagueDashboard) {
      await loadSession();
    } else {
      logout();
    }
  };
}

function buildSquadSwitcher(data) {
  squadSwitcher.innerHTML = "";

  if (!data.opposition) {
    squadSwitcher.classList.add("hidden");
    return;
  }

  squadSwitcher.classList.remove("hidden");

  const ownButton = document.createElement("button");
  ownButton.className = "squad-tab active";
  ownButton.textContent = "Your Squad";

  const oppositionButton = document.createElement("button");
  oppositionButton.className = "squad-tab";
  oppositionButton.textContent = data.opposition.team.name;

  ownButton.addEventListener("click", () => {
    activeSquad = "own";
    currentPlayers = ownPlayers;
    playerCount.textContent = ownPlayers.length + " Approved Players";
    search.value = "";
    ownButton.classList.add("active");
    oppositionButton.classList.remove("active");
    renderPlayers(currentPlayers);
  });

  oppositionButton.addEventListener("click", () => {
    activeSquad = "opposition";
    currentPlayers = oppositionPlayers;
    playerCount.textContent =
      oppositionPlayers.length + " Opposition Players";
    search.value = "";
    oppositionButton.classList.add("active");
    ownButton.classList.remove("active");
    renderPlayers(currentPlayers);
  });

  squadSwitcher.appendChild(ownButton);
  squadSwitcher.appendChild(oppositionButton);
}

function logout() {
  localStorage.removeItem("mfflSessionId");
  dashboardData = null;
  currentPlayers = [];
  ownPlayers = [];
  oppositionPlayers = [];
  playersDiv.innerHTML = "";
  teamIdInput.value = "";
  adminCodeInput.value = "";
  loginError.textContent = "";
  showLogin();
}

function renderPlayers(players) {
  playersDiv.innerHTML = "";

  if (!players.length) {
    playersDiv.innerHTML = `<p class="empty-message">No approved players found.</p>`;
    return;
  }

  players.forEach(player => {
    const card = document.createElement("div");
    card.className = "player-card";

    const bmi = player.showBMI
      ? `<div class="player-bmi"><strong>BMI:</strong> ${player.bmi}</div>`
      : "";

    const note = player.note
      ? `<div class="league-note"><strong>League Note</strong>${player.note}</div>`
      : "";

    const photoStyle = player.photo
      ? `background-image:url('${player.photo}')`
      : "";

    card.innerHTML = `
      <div class="player-photo" style="${photoStyle}"></div>

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