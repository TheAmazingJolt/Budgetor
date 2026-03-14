FROM node:20-alpine
WORKDIR /app
RUN corepack enable && corepack prepare pnpm@10.26.1 --activate
COPY . .
RUN pnpm install --no-frozen-lockfile && pnpm --filter @workspace/api-server build
EXPOSE 8080
CMD ["node", "artifacts/api-server/dist/index.cjs"]