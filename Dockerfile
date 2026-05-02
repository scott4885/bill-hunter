# Multi-stage build: compile TypeScript in builder, ship dist + prod deps only.

FROM node:20-alpine AS build
WORKDIR /app
COPY package*.json tsconfig.json ./
RUN npm install --no-audit --no-fund
COPY src ./src
RUN npm run build

FROM node:20-alpine AS runtime
WORKDIR /app
ENV NODE_ENV=production
ENV PORT=8080
RUN apk add --no-cache curl
COPY package*.json ./
RUN npm install --no-audit --no-fund --omit=dev --ignore-scripts && npm cache clean --force
COPY --from=build /app/dist ./dist

EXPOSE 8080
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
  CMD wget --quiet --tries=1 --spider http://localhost:8080/healthz || exit 1

CMD ["node", "dist/server.js"]
