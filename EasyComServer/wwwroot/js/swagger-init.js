document.getElementById('server-url-display').textContent =
  location.protocol + '//' + location.host + '/api/v1';

SwaggerUIBundle({
  url: './openapi.yaml',
  dom_id: '#swagger-ui',
  presets: [SwaggerUIBundle.presets.apis, SwaggerUIBundle.SwaggerUIStandalonePreset],
  layout: 'BaseLayout',
  deepLinking: true,
  tryItOutEnabled: true,
  requestInterceptor: req => {
    req.url = req.url.replace(/^http:\/\/localhost:\d+/, location.protocol + '//' + location.host);
    return req;
  },
});
