FROM node:22-alpine AS dependencies
WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci

FROM node:22-alpine AS build
WORKDIR /app
COPY --from=dependencies /app/node_modules ./node_modules
COPY . .
RUN npm run build

FROM node:22-alpine AS runtime
ENV HOST=0.0.0.0
ENV PORT=4321
ENV NODE_ENV=production
WORKDIR /app
RUN addgroup -S astro && adduser -S astro -G astro
COPY --from=build --chown=astro:astro /app/dist ./dist
COPY --from=build --chown=astro:astro /app/package.json ./package.json
COPY --from=dependencies --chown=astro:astro /app/node_modules ./node_modules
USER astro
EXPOSE 4321
CMD ["node", "./dist/server/entry.mjs"]
