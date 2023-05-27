FROM nginx:stable-alpine

COPY /build /usr/share/nginx/html

COPY ./manifest.xml /tmp/manifest.xml

RUN sed -i 's|https://localhost:3000/|https://outlooktranslate.wiltechs.com/|g' /tmp/manifest.xml

RUN cp /tmp/manifest.xml /usr/share/nginx/html/manifest.xml

RUN  sed -i '12a error_page 404 /index.html;' /etc/nginx/conf.d/default.conf

EXPOSE 80
CMD ["nginx", "-g", "daemon off;"]
