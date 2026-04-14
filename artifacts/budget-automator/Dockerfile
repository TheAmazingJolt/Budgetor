FROM node:20-slim
WORKDIR /app
RUN corepack enable && corepack prepare pnpm@10.26.1 --activate
COPY . .
ENV PORT=3000
ENV BASE_PATH=/
ARG VITE_API_BASE_URL
ENV VITE_API_BASE_URL=$VITE_API_BASE_URL
RUN rm -rf artifacts/budget-automator/dist
ARG BUILD_TIME
ENV VITE_BUILD_TIME=${BUILD_TIME}
RUN BUILD_TIME=$(date -u +"%Y-%m-%dT%H:%M:%SZ") && echo "Build time: $BUILD_TIME" && NODE_ENV=production VITE_BUILD_TIME=$BUILD_TIME pnpm install --frozen-lockfile && pnpm --filter @workspace/budget-automator build
EXPOSE 3000
# We point to the actual path inside the monorepo
CMD ["sh", "-c", "cd artifacts/budget-automator && pnpm vite preview --host 0.0.0.0 --port ${PORT:-3000}"]


