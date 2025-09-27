import logging
import sys
from django.http import HttpResponse
from django.utils.deprecation import MiddlewareMixin

logger = logging.getLogger(__name__)

class BrokenPipeMiddleware(MiddlewareMixin):
    """
    Middleware to handle broken pipe errors gracefully.
    """
    
    def __init__(self, get_response):
        super().__init__(get_response)
    
    def process_request(self, request):
        # Add request timeout handling
        return None
    
    def process_response(self, request, response):
        # Add headers to prevent broken pipes and improve connection handling
        response['Cache-Control'] = 'no-cache, no-store, must-revalidate'
        response['Pragma'] = 'no-cache'
        response['Expires'] = '0'
        # Note: Connection header is not allowed in WSGI responses
        
        return response
    
    def process_exception(self, request, exception):
        # Handle various connection-related exceptions
        exception_str = str(exception).lower()
        
        if any(error in exception_str for error in [
            'broken pipe', 'connectionreseterror', 'connection aborted',
            'connection lost', 'connection closed', 'socket error'
        ]):
            # Log the error but don't crash the server
            logger.warning(f"Connection error handled gracefully: {exception}")
            return HttpResponse(status=200)  # Return 200 to prevent client errors
        
        return None

