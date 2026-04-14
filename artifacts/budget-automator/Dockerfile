FROM node:20-slim
WORKDIR /app
RUN corepack enable && corepack prepare pnpm@10.26.1 --activate
COPY . .
ENV PORT=3000
ENV BASE_PATH=/
ARG VITE_API_BASE_URL
ENV VITE_API_BASE_URL=$VITE_API_BASE_URL
RUN NODE_ENV=production pnpm install --frozen-lockfile && pnpm --filter @workspace/budget-automator build
EXPOSE 3000
# We point to the actual path inside the monorepo
CMD ["sh", "-c", "pnpm --filter @workspace/budget-automator serve --dir artifacts/budget-automator/dist"]

