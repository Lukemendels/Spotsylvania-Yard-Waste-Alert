const statusText = {
  accepting: ["Accepting yard waste", "Chancellor is accepting residential yard waste."],
  filling: ["Yard waste area filling up", "Chancellor is still accepting yard waste. Check again before leaving home."],
  full: ["Yard waste area full", "Chancellor has paused yard waste acceptance. Check back before bringing a load."],
  unavailable: ["Current status unavailable", "Check the county’s Yard Waste page before bringing a load."]
};

const notice = document.querySelector("#website-alert");
const title = document.querySelector("#alert-title");
const message = document.querySelector("#alert-message");
const confirmation = document.querySelector("#alert-time");
const eastern = new Intl.DateTimeFormat("en-US", {
  timeZone: "America/New_York", month: "short", day: "numeric",
  hour: "numeric", minute: "2-digit", timeZoneName: "short"
});

let latestVersion = 0;
let expiration;
let requestInFlight = false;

function readStatus(feedUrl) {
  return new Promise((resolve, reject) => {
    const random = crypto.randomUUID().replaceAll("-", "");
    const callback = "yardWasteStatus_" + random;
    const url = new URL(feedUrl);
    url.searchParams.set("feed", "1");
    url.searchParams.set("callback", callback);
    const script = document.createElement("script");
    const timeout = setTimeout(() => finish(new Error("Timed out")), 8000);

    function finish(error, status) {
      clearTimeout(timeout);
      delete window[callback];
      script.remove();
      if (error) reject(error);
      else resolve(status);
    }

    window[callback] = status => finish(null, status);
    script.onerror = () => finish(new Error("Unable to load status"));
    script.src = url.href;
    document.head.append(script);
  });
}

function render(status, confirmedAt, expiresAt) {
  clearTimeout(expiration);
  const [heading, detail] = statusText[status];
  if (title.textContent !== heading) title.textContent = heading;
  if (message.textContent !== detail) message.textContent = detail;
  notice.dataset.status = status;
  confirmation.hidden = status === "unavailable";
  if (!confirmation.hidden) {
    confirmation.dateTime = confirmedAt;
    const text = "Confirmed " + eastern.format(new Date(confirmedAt));
    if (confirmation.textContent !== text) confirmation.textContent = text;
    expiration = setTimeout(() => render("unavailable"), expiresAt - Date.now());
  }
}

async function refreshStatus(feedUrl) {
  if (requestInFlight) return;
  requestInFlight = true;
  try {
    const data = await readStatus(feedUrl);
    if (!data || !Number.isInteger(data.version) || data.version < 0) {
      render("unavailable");
      return;
    }
    if (data.version < latestVersion) return;
    latestVersion = data.version;

    const confirmed = Date.parse(data.confirmedAt);
    const expiry = Date.parse(data.expiresAt);
    if (data.status !== "unavailable" && Object.hasOwn(statusText, data.status) &&
        Number.isFinite(confirmed) && confirmed <= Date.now() &&
        Number.isFinite(expiry) && expiry > Date.now() && expiry > confirmed) {
      render(data.status, data.confirmedAt, expiry);
    } else {
      render("unavailable");
    }
  } catch {
    render("unavailable");
  } finally {
    requestInFlight = false;
  }
}

const feedUrl = window.YARD_WASTE_FEED_URL;
render("unavailable");
if (/^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(feedUrl)) {
  refreshStatus(feedUrl);
  setInterval(() => {
    if (!document.hidden) refreshStatus(feedUrl);
  }, 15000);
  document.addEventListener("visibilitychange", () => {
    if (!document.hidden) refreshStatus(feedUrl);
  });
}
