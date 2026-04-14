FROM node:20-slim
WORKDIR /app
RUN corepack enable && corepack prepare pnpm@10.33.0 --activate
COPY . .
RUN rm -rf artifacts/api-server/dist artifacts/budget-automator/dist
RUN echo "=== API BUILD START ===" && \
    pnpm install --no-lockfile && \
    pnpm --filter @workspace/api-server build && \
    echo "=== API BUILD COMPLETE ==="
RUN date -u +"%Y-%m-%dT%H:%M:%SZ" > /app/BUILD_TIME
EXPOSE 8080
CMD ["node", "artifacts/api-server/dist/index.cjs"]
