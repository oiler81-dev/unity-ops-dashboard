function json(status, body) {
  return {
    status,
    headers: {
      "Content-Type": "application/json"
    },
    body
  };
}

function ok(body) {
  return json(200, body);
}

function badRequest(message, extra = {}) {
  return json(400, {
    ok: false,
    error: message,
    ...extra
  });
}

function forbidden(message = "Forbidden") {
  return json(403, {
    ok: false,
    error: message
  });
}

function serverError(error, message = "Server error") {
  // Intentionally omit error.message / stack from the client response so we
  // don't leak implementation details. Callers are responsible for logging
  // the raw error (typically context.log.error(...)).
  return json(500, {
    ok: false,
    error: message
  });
}

module.exports = {
  json,
  ok,
  badRequest,
  forbidden,
  serverError
};
