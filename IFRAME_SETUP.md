# 📧 Email Processor - Iframe Integration Guide

Transform your email processing system into an embeddable iframe that can be integrated into any web application or software!

## 🚀 Quick Start

### 1. Start the Server
```bash
python main.py
```

### 2. Access the Iframe
Visit: `http://localhost:8000/iframe`

### 3. Get Embed Code
Visit: `http://localhost:8000/embed-code` for copy-paste iframe code

### 4. See Integration Example
Open: `example_embed.html` in your browser

## 🔧 Basic Integration

### Simple Iframe Embed
```html
<iframe 
    src="http://localhost:8000/iframe" 
    width="100%" 
    height="800" 
    frameborder="0" 
    style="border: 1px solid #ddd; border-radius: 8px; box-shadow: 0 4px 8px rgba(0,0,0,0.1);"
    allowfullscreen>
</iframe>
```

### Responsive Integration
```html
<div style="position: relative; width: 100%; height: 0; padding-bottom: 56.25%;">
    <iframe 
        src="http://localhost:8000/iframe" 
        style="position: absolute; top: 0; left: 0; width: 100%; height: 100%; border: none;"
        allowfullscreen>
    </iframe>
</div>
```

## 🎯 Available Endpoints

| Endpoint | Description |
|----------|-------------|
| `/iframe` | Main iframe interface |
| `/embed-code` | Get iframe embed code |
| `/results` | Get recent extraction results |
| `/stats` | Get processing statistics |
| `/process-now` | Trigger manual processing |
| `/health` | Health check |

## ✨ Features

### 🎮 User Interface
- **Modern Design**: Clean, responsive interface with gradient backgrounds
- **Real-time Updates**: Auto-refreshing statistics every 30 seconds
- **Interactive Controls**: Process emails, view results, download data
- **Status Indicators**: Live connection and processing status
- **Notifications**: Toast notifications for all actions

### 📊 Functionality
- **Manual Email Processing**: Trigger processing with one click
- **Live Statistics**: Total processed, success rate, confidence scores
- **Results Viewer**: Display recent extraction results with details
- **Data Export**: Download statistics and results as JSON
- **Health Monitoring**: System status and uptime tracking

### 📱 Responsive Design
- **Mobile Friendly**: Works on phones, tablets, and desktops
- **Flexible Sizing**: Adapts to any iframe dimensions
- **Touch Optimized**: Touch-friendly controls and interactions

## 🔒 Security Features

### Iframe-Optimized Headers
- `X-Frame-Options: ALLOWALL` - Allows iframe embedding from any domain
- `Content-Security-Policy: frame-ancestors *` - Permits framing
- Proper CORS configuration for cross-origin requests

### Production Security
For production deployment, consider:
- Restricting `X-Frame-Options` to specific domains
- Using HTTPS for all communications
- Implementing proper authentication if needed
- Setting up rate limiting

## 🛠️ Customization

### Styling
The iframe interface uses CSS custom properties that can be overridden:

```css
iframe {
    width: 100%;
    height: 600px;
    border: 2px solid #your-brand-color;
    border-radius: 12px;
    box-shadow: 0 8px 24px rgba(0,0,0,0.15);
}
```

### Size Configurations
- **Minimum recommended height**: 500px
- **Optimal height**: 800px
- **Mobile height**: 600px
- **Width**: Always use 100% for responsiveness

## 🌐 Integration Examples

### React Component
```jsx
import React from 'react';

const EmailProcessorIframe = ({ height = "800px" }) => {
    return (
        <iframe
            src="http://localhost:8000/iframe"
            width="100%"
            height={height}
            style={{
                border: '1px solid #e2e8f0',
                borderRadius: '8px',
                boxShadow: '0 4px 8px rgba(0,0,0,0.1)'
            }}
            allowFullScreen
        />
    );
};

export default EmailProcessorIframe;
```

### Vue Component
```vue
<template>
    <iframe
        :src="iframeUrl"
        width="100%"
        :height="height"
        frameborder="0"
        allowfullscreen
        class="email-processor-iframe"
    />
</template>

<script>
export default {
    name: 'EmailProcessorIframe',
    props: {
        height: {
            type: String,
            default: '800px'
        }
    },
    data() {
        return {
            iframeUrl: 'http://localhost:8000/iframe'
        }
    }
}
</script>

<style scoped>
.email-processor-iframe {
    border: 1px solid #e2e8f0;
    border-radius: 8px;
    box-shadow: 0 4px 8px rgba(0,0,0,0.1);
}
</style>
```

### WordPress Integration
```php
// Add to functions.php
function email_processor_shortcode($atts) {
    $atts = shortcode_atts(array(
        'height' => '800',
        'width' => '100%'
    ), $atts);
    
    return '<iframe 
        src="http://localhost:8000/iframe" 
        width="' . $atts['width'] . '" 
        height="' . $atts['height'] . '" 
        frameborder="0" 
        style="border: 1px solid #ddd; border-radius: 8px;" 
        allowfullscreen>
    </iframe>';
}
add_shortcode('email_processor', 'email_processor_shortcode');

// Usage: [email_processor height="600"]
```

## 🚀 Production Deployment

### 1. Update Configuration
```python
# config.py
API_HOST = "0.0.0.0"  # Allow external connections
API_PORT = 8000
```

### 2. Environment Variables
```bash
# Update URLs for production
export API_HOST=yourdomain.com
export API_PORT=443  # For HTTPS
```

### 3. HTTPS Setup
```python
# For production with SSL
if __name__ == "__main__":
    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=443,
        ssl_keyfile="/path/to/key.pem",
        ssl_certfile="/path/to/cert.pem"
    )
```

### 4. Update Iframe URLs
Replace `http://localhost:8000` with your production URL:
```html
<iframe src="https://yourdomain.com/iframe" ...>
```

## 🔧 Troubleshooting

### Common Issues

**Iframe Not Loading**
- Check if the server is running on the correct port
- Verify CORS settings allow your domain
- Check browser console for security errors

**Authentication Issues**
- Ensure environment variables are properly set
- Check Microsoft Graph API credentials
- Verify Gemini AI API key is valid

**Styling Issues**
- Check iframe dimensions are adequate (min 500px height)
- Verify CSS is not conflicting with iframe content
- Test on different browsers and devices

### Debug Mode
Enable debug logging:
```python
import logging
logging.basicConfig(level=logging.DEBUG)
```

## 📞 Support

- **API Documentation**: Visit `/docs` for FastAPI auto-generated docs
- **Health Check**: Visit `/health` to verify system status
- **Embed Code**: Visit `/embed-code` for copy-paste iframe code

## 🎉 Integration Success!

Once integrated, you'll have:
- ✅ Full email processing functionality in your app
- ✅ Real-time statistics and monitoring
- ✅ Modern, responsive user interface
- ✅ Secure, iframe-optimized architecture
- ✅ Easy maintenance and updates

Your email processing system is now ready to be embedded anywhere! 🚀 