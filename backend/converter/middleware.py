import logging
from django.http import HttpResponse
from django.utils.deprecation import MiddlewareMixin

logger = logging.getLogger(__name__)

class BrokenPipeMiddleware(MiddlewareMixin):
    """
    Middleware to handle broken pipe errors gracefully.
    """
    
    def process_response(self, request, response):
        # Add headers to prevent broken pipes
        response['Connection'] = 'keep-alive'
        response['Cache-Control'] = 'no-cache, no-store, must-revalidate'
        response['Pragma'] = 'no-cache'
        response['Expires'] = '0'
        
        return response
    
    def process_exception(self, request, exception):
        # Log broken pipe errors but don't crash the server
        if 'Broken pipe' in str(exception) or 'ConnectionResetError' in str(exception):
            logger.warning(f"Broken pipe error handled: {exception}")
            return HttpResponse(status=200)  # Return 200 to prevent client errors
        
        return None

