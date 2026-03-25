export interface GraphWebhookNotification {
  subscriptionId: string;
  clientState?: string;
  changeType?: string;
  resource?: string;
  tenantId?: string;
  lifecycleEvent?: string;
  resourceData?: { id?: string };
}

export interface GraphWebhookEnvelope {
  value: GraphWebhookNotification[];
}

export function getGraphValidationToken(request: Request): string | null {
  return new URL(request.url).searchParams.get("validationToken");
}

export function parseGraphWebhookBody(body: string): GraphWebhookEnvelope {
  const parsed = JSON.parse(body) as GraphWebhookEnvelope;
  return {
    value: Array.isArray(parsed.value) ? parsed.value : [],
  };
}

export function isGraphClientStateValid(
  notification: GraphWebhookNotification,
  expectedClientState: string,
): boolean {
  return notification.clientState === expectedClientState;
}
