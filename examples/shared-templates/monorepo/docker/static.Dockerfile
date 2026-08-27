ARG RUNTIME_IMAGE=registry.buluttakin.com/nginx:1.27-alpine
FROM ${RUNTIME_IMAGE}

COPY current/ /srv/monorepo/current/
COPY nginx/default.conf /etc/nginx/conf.d/default.conf

RUN nginx -t

