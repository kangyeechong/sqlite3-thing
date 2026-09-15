"""
Runs the development web server.

NOT for production use - Flask's built-in server is single-threaded
and not hardened, and there's no HTTPS or login rate-limiting yet (see
the Step 4 summary for the full list). Fine for trying the tool out
locally; a real deployment needs a proper WSGI server behind HTTPS,
which is a hosting decision for later.

Usage:
    python3 run_server.py --db ledger.db
    (then open http://127.0.0.1:5000 in a browser)
"""

import argparse

from app.web import create_app

_LOOPBACK_HOSTS = {"127.0.0.1", "localhost"}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--db", default="ledger.db")
    parser.add_argument("--port", type=int, default=5000)
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument(
        "--debug", action="store_true",
        help="Enable Flask's interactive debugger. Only ever safe on 127.0.0.1/localhost - "
             "the debugger allows arbitrary code execution from anyone who can reach it.",
    )
    args = parser.parse_args()

    if args.debug and args.host not in _LOOPBACK_HOSTS:
        raise SystemExit(
            f"Refusing to start with --debug on host '{args.host}'. "
            f"The Flask debugger allows arbitrary code execution from "
            f"anyone who can reach it - only use --debug with the "
            f"default 127.0.0.1/localhost host."
        )

    app = create_app(args.db)
    app.run(host=args.host, port=args.port, debug=args.debug)


if __name__ == "__main__":
    main()
