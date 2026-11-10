"""
Main entry point for the outlook_mcp_server package when executed as a module.
This allows the package to be run with 'python -m outlook_mcp_server'.
"""

import sys
from fastmcp import FastMCP

# Import from the package
try:
    # When running as module
    from .backend.outlook_session.session_manager import OutlookSessionManager
    from .tools.registration import register_all_tools
    from .remote import RemoteConfigError, build_token_verifier, load_settings
    from .bridge import bridge_settings_from_env, start_bridge_thread
except ImportError:
    # When running as script or direct execution
    from outlook_mcp_server.backend.outlook_session.session_manager import OutlookSessionManager
    from outlook_mcp_server.tools.registration import register_all_tools
    from outlook_mcp_server.remote import RemoteConfigError, build_token_verifier, load_settings
    from outlook_mcp_server.bridge import bridge_settings_from_env, start_bridge_thread

def test_outlook_connection() -> bool:
    """Test Outlook connection before starting the server.
    
    Returns:
        bool: True if connection successful, False otherwise
    """
    try:
        with OutlookSessionManager() as session:
            # Test basic folder access
            inbox = session.get_folder()
            if inbox and hasattr(inbox, 'Name'):
                return True
    except Exception as e:
        print(f"Outlook connection test failed: {str(e)}", file=sys.stderr)
        return False
    return False

def main():
    """Main function to start the Outlook MCP Server.
    
    This function serves as the entry point for module execution.
    It tests the Outlook connection before starting the MCP server.

    Transport is stdio by default. With ``--transport http`` (or the
    OUTLOOK_MCP_TRANSPORT environment variable) the server listens on the
    loopback interface and requires a bearer token on every request. See
    docs/remote-access.md.
    """
    try:
        settings = load_settings()
        bridge_settings = bridge_settings_from_env() if settings.is_http else None
    except RemoteConfigError as e:
        print(f"Configuration error: {e}", file=sys.stderr)
        sys.exit(2)

    try:
        # Test Outlook connection first
        if not test_outlook_connection():
            print("Error: Unable to connect to Outlook. Please ensure Outlook is installed and running.", file=sys.stderr)
            sys.exit(1)
        
        print("Outlook connection successful. Starting MCP server...", file=sys.stderr)
        
        # Initialize FastMCP server, with bearer-token auth in HTTP mode
        auth = build_token_verifier(settings.token) if settings.is_http else None
        mcp = FastMCP("outlook-assistant", auth=auth)
        
        # Register all MCP tools
        register_all_tools(mcp)
        
        # Run the MCP server
        if settings.is_http:
            print(
                f"Listening on http://{settings.host}:{settings.port}{settings.path} "
                "(bearer token required)",
                file=sys.stderr,
            )
            if bridge_settings is not None:
                print(f"Starting outbound bridge to {bridge_settings.url}", file=sys.stderr)
                start_bridge_thread(bridge_settings)
            # Stateless JSON responses: every request is self-contained, which is what the
            # bridge/relay path expects and what tool-only clients need.
            mcp.run(
                transport="http", host=settings.host, port=settings.port, path=settings.path,
                stateless_http=True, json_response=True,
            )
        else:
            mcp.run()
        
    except KeyboardInterrupt:
        print("\nMCP server stopped by user.", file=sys.stderr)
        sys.exit(0)
    except Exception as e:
        print(f"Error starting server: {str(e)}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()