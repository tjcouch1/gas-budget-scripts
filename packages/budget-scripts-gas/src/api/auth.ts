/**
 * API key authentication (web-plan.md section 4.4, AGENTS.md section 1). The key lives only in
 * Script Properties and travels only in the JSON body of POST requests.
 */

/** Script Property that holds the shared API key */
export const API_KEY_PROPERTY = "API_KEY";

/**
 * Compares two strings in time that depends only on their lengths, not on where they differ
 */
export function constantTimeEquals(a: string, b: string): boolean {
  const length = Math.max(a.length, b.length);
  let diff = a.length ^ b.length;
  for (let i = 0; i < length; i++)
    diff |=
      (i < a.length ? a.charCodeAt(i) : 0) ^
      (i < b.length ? b.charCodeAt(i) : 0);
  return diff === 0;
}

/**
 * Whether the request's API key matches the `API_KEY` Script Property. Always false if the
 * property is unset or empty
 */
export function isAuthorized(apiKey: unknown): boolean {
  const expected =
    PropertiesService.getScriptProperties().getProperty(API_KEY_PROPERTY);
  if (!expected) return false;
  if (typeof apiKey !== "string" || apiKey === "") return false;
  return constantTimeEquals(apiKey, expected);
}

/**
 * Generates a random API key, stores it in Script Properties as `API_KEY`, and logs it once
 *
 * @param overwrite set to true to replace an existing key. Otherwise throws if one exists
 */
export function generateApiKey(overwrite: boolean) {
  const properties = PropertiesService.getScriptProperties();
  if (properties.getProperty(API_KEY_PROPERTY) && !overwrite)
    throw new Error(
      `${API_KEY_PROPERTY} already exists in Script Properties. To rotate it, delete the ${API_KEY_PROPERTY} property in Project Settings > Script Properties, then run setupGenerateApiKey again.`,
    );

  const apiKey = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, "");
  properties.setProperty(API_KEY_PROPERTY, apiKey);
  Logger.log(
    `Stored a new ${API_KEY_PROPERTY} in Script Properties. Copy it into the web app's Settings dialog: ${apiKey}`,
  );
}

/**
 * Run once from the Apps Script editor: creates the `API_KEY` Script Property for the web app.
 * Refuses to overwrite an existing key (delete the property first to rotate it). Takes no
 * parameters so a trigger event can never be mistaken for an overwrite flag
 */
export function setupGenerateApiKey() {
  generateApiKey(false);
}
