"""
WSGI config for excel_backend project.

It exposes the WSGI callable as a module-level variable named ``application``.

For more information on this file, see
https://docs.djangoproject.com/en/5.2/howto/deployment/wsgi/
"""

import os
import signal
import logging

from django.core.wsgi import get_wsgi_application

# Suppress broken pipe errors at the system level
signal.signal(signal.SIGPIPE, signal.SIG_DFL)

# Configure logging to reduce broken pipe noise
logging.getLogger('django.server').setLevel(logging.WARNING)

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'excel_backend.settings')

# Get the Django WSGI application
django_app = get_wsgi_application()

def application(environ, start_response):
    """
    WSGI application wrapper to handle broken pipe errors gracefully.
    """
    try:
        return django_app(environ, start_response)
    except (BrokenPipeError, ConnectionResetError, OSError) as e:
        # Handle broken pipe and connection errors gracefully
        if 'Broken pipe' in str(e) or 'Connection reset' in str(e):
            # Return a minimal response to prevent server crashes
            start_response('200 OK', [('Content-Type', 'text/plain')])
            return [b'Connection closed']
        else:
            # Re-raise other errors
            raise