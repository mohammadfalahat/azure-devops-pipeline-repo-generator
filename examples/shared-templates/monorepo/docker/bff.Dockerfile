ARG RUNTIME_IMAGE=registry.buluttakin.com/node:20-alpine
FROM ${RUNTIME_IMAGE}

ARG BFF_ENTRY=main.js
ENV BFF_ENTRY=${BFF_ENTRY}

WORKDIR /app
COPY app/ /app/

CMD ["/bin/sh", "-ec", "exec node \"$BFF_ENTRY\""]

