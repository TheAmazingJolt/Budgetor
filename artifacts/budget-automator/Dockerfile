FROM node:20-slim
WORKDIR /app
RUN corepack enable && corepack prepare pnpm@10.33.0 --activate
COPY . .
ENV PORT=3000
ARG VITE_API_BASE_URL
ENV VITE_API_BASE_URL=$VITE_API_BASE_URL
RUN rm -rf artifacts/budget-automator/dist artifacts/api-server/dist
RUN export VITE_BUILD_TIME=$(date -u +"%Y-%m-%dT%H:%M:%SZ") && \
    echo "=== BUILD START: $VITE_BUILD_TIME ===" && \
    pnpm install --no-frozen-lockfile && \
    pnpm --filter @workspace/budget-automator build && \
    echo "=== BUILD COMPLETE ==="
EXPOSE 3000
CMD ["sh", "-c", "cd artifacts/budget-automator && pnpm vite preview --host 0.0.0.0 --port ${PORT:-3000}"]
