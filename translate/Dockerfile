FROM nginx:stable-alpine
COPY /build /usr/share/nginx/html

COPY ./manifest.xml /app/manifest.xml
RUN sed -i 's|https://localhost:3000/|https://outlooktranslateplugin.com/|g' /app/manifest.xml


RUN  sed -i '12a error_page 404 /index.html;' /etc/nginx/conf.d/default.conf

EXPOSE 80
CMD ["nginx", "-g", "daemon off;"]
