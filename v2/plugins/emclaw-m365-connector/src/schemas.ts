import { Type } from "typebox";

export const workflowIdSchema = Type.String({ description: "Workflow identifier, not a path." });
export const callerSchema = Type.String({ description: "IOC limitation: explicit caller user_ref, validated against roster role/status." });
export const specSchema = Type.String({ description: "Workflow spec JSON string. For member tools this is the backend workflow body; product-facing skills must also pass explicit app/service/operation metadata." });

export const normalizedCreateParams = {
  spec_json: specSchema,
  app_id: Type.Optional(Type.String({ description: "Registered app id, for example microsoft365." })),
  service: Type.Optional(Type.String({ description: "Service area, for example calendar." })),
  operation: Type.Optional(Type.String({ description: "Explicit operation template, for example calendar.metadata_daily_status_check." })),
  operation_template: Type.Optional(Type.String({ description: "Alias for operation when a product skill names a template." })),
  scope: Type.Optional(Type.String({ description: "Workflow scope. Member tools only allow single_user." })),
  schedule: Type.Optional(Type.Any()),
  timezone: Type.Optional(Type.String()),
  live_effects: Type.Optional(Type.Boolean()),
  delivery: Type.Optional(Type.Any()),
  content_access: Type.Optional(Type.Any()),
};
