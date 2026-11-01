import { describe, it, expect, vi, beforeEach } from "vitest";

// Mock the client so we can capture exactly what reaches core.write (the seam
// right before the HTTP request). No live TeamSnap calls are made.
const writeMock = vi.fn();

vi.mock("../../src/api/client.js", () => ({
  teamsnapClient: {
    isAuthenticated: () => true,
    reloadCredentials: () => {},
    getCore: () => ({ write: writeMock }),
  },
}));

const { handleCreateEvent, handleUpdateEvent } = await import("../../src/tools/handlers/events.js");

function parse(result: { content: Array<{ type: string; text?: string }> }) {
  const c = result.content[0];
  return JSON.parse(c.type === "text" ? (c.text as string) : "{}");
}

const baseCreate = {
  team_id: "123",
  name: "Game vs Hawks",
  start_date: "2026-11-08T15:00:00Z",
  preview: false,
};

beforeEach(() => {
  writeMock.mockReset();
  writeMock.mockResolvedValue({ id: 999, name: "Game vs Hawks", start_date: "2026-11-08T15:00:00Z" });
});

describe("teamsnap_create_event notify_team", () => {
  it("sends notify_team and notify_team_as_member_id in the POST body when passed", async () => {
    const r = await handleCreateEvent({
      ...baseCreate,
      notify_team: true,
      notify_team_as_member_id: 4242,
      label: "Semifinal",
    });
    expect(r.isError).toBeFalsy();
    expect(writeMock).toHaveBeenCalledTimes(1);
    const [method, , body] = writeMock.mock.calls[0];
    expect(method).toBe("POST");
    expect(body.notify_team).toBe(true);
    expect(body.notify_team_as_member_id).toBe(4242);
    expect(body.label).toBe("Semifinal");
  });

  it("omits notify fields from the POST body when not passed", async () => {
    await handleCreateEvent({ ...baseCreate });
    expect(writeMock).toHaveBeenCalledTimes(1);
    const body = writeMock.mock.calls[0][2];
    expect(body).not.toHaveProperty("notify_team");
    expect(body).not.toHaveProperty("notify_team_as_member_id");
    expect(body).not.toHaveProperty("label");
  });

  it("passes notify_team: false through when explicitly set", async () => {
    await handleCreateEvent({ ...baseCreate, notify_team: false });
    expect(writeMock.mock.calls[0][2].notify_team).toBe(false);
  });

  it("rejects a non-boolean notify_team without writing", async () => {
    const r = await handleCreateEvent({ ...baseCreate, notify_team: "yes" });
    expect(r.isError).toBe(true);
    expect(writeMock).not.toHaveBeenCalled();
  });

  it("preview shows notify fields and does not write", async () => {
    const on = parse(await handleCreateEvent({ ...baseCreate, preview: true, notify_team: true }));
    expect(on.notify_team).toBe(true);
    expect(on.template.data).toContainEqual({ name: "notify_team", value: true });
    const off = parse(await handleCreateEvent({ ...baseCreate, preview: true }));
    expect(off.notify_team).toBe(false);
    expect(off.notification).toMatch(/No notification/);
    expect(writeMock).not.toHaveBeenCalled();
  });
});

describe("teamsnap_update_event notify_team + scores", () => {
  it("merges top-level notify_team into the PATCH body", async () => {
    const r = await handleUpdateEvent({
      event_id: "555",
      patch: { name: "Moved" },
      notify_team: true,
      notify_team_as_member_id: "4242",
      preview: false,
    });
    expect(r.isError).toBeFalsy();
    expect(writeMock).toHaveBeenCalledTimes(1);
    const [method, endpoint, body] = writeMock.mock.calls[0];
    expect(method).toBe("PATCH");
    expect(String(endpoint)).toContain("555");
    expect(body).toEqual({ name: "Moved", notify_team: true, notify_team_as_member_id: "4242" });
  });

  it("omits notify fields from the PATCH body when not passed", async () => {
    await handleUpdateEvent({ event_id: "555", patch: { name: "Moved" }, preview: false });
    const body = writeMock.mock.calls[0][2];
    expect(body).toEqual({ name: "Moved" });
    expect(body).not.toHaveProperty("notify_team");
  });

  it("refuses notify_team inside patch", async () => {
    const r = await handleUpdateEvent({ event_id: "555", patch: { notify_team: true }, preview: false });
    expect(r.isError).toBe(true);
    expect(writeMock).not.toHaveBeenCalled();
  });

  it("allows points_for_team, points_for_opponent and label through the allowlist", async () => {
    const r = await handleUpdateEvent({
      event_id: "555",
      patch: { points_for_team: 4, points_for_opponent: 2, label: "Final" },
      preview: false,
    });
    expect(r.isError).toBeFalsy();
    expect(writeMock.mock.calls[0][2]).toEqual({ points_for_team: 4, points_for_opponent: 2, label: "Final" });
  });

  it("still rejects unknown patch keys", async () => {
    const r = await handleUpdateEvent({ event_id: "555", patch: { bogus: 1 }, preview: false });
    expect(r.isError).toBe(true);
    expect(writeMock).not.toHaveBeenCalled();
  });

  it("preview shows the merged body and notify state without writing", async () => {
    const p = parse(
      await handleUpdateEvent({ event_id: "555", patch: { points_for_team: 3 }, notify_team: true })
    );
    expect(p.preview).toBe(true);
    expect(p.with).toEqual({ points_for_team: 3, notify_team: true });
    expect(p.notify_team).toBe(true);
    expect(writeMock).not.toHaveBeenCalled();
  });
});
