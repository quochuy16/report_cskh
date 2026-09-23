// POST /api/logout
import {api, send} from "./_lib/http.js";
import {clearSession} from "./_lib/auth.js";

export default api(["POST"], async (req, res) => { clearSession(req, res); send(res, 200, {ok: true}); });
