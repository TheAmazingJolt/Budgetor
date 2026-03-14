FROM node:20-alpine
WORKDIR /app
RUN corepack enable && corepack prepare pnpm@10.26.1 --activate
COPY . .
RUN PORT=3000 BASE_PATH=/ NODE_ENV=production pnpm install --no-lockfile && pnpm --filter @workspace/budget-automator build
EXPOSE 3000
CMD ["sh", "-c", "PORT=${PORT:-3000} BASE_PATH=/ pnpm --filter @workspace/budget-automator serve"]
