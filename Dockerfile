FROM node:20-slim
WORKDIR /app
RUN corepack enable && corepack prepare pnpm@10.26.1 --activate
COPY . .
RUN pnpm --version
RUN cat pnpm-workspace.yaml
RUN pnpm install --no-lockfile && pnpm --filter @workspace/api-server build
RUN date -u +"%Y-%m-%dT%H:%M:%SZ" > /app/BUILD_TIME
EXPOSE 8080
CMD ["node", "artifacts/api-server/dist/index.cjs"]