FROM node:20-alpine
WORKDIR /app
COPY . .
RUN npm install -g pnpm && pnpm install && pnpm --filter @workspace/api-server build
EXPOSE 8080
CMD ["node", "artifacts/api-server/dist/index.cjs"]