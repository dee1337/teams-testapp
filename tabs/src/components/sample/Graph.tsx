import React from "react";
import { Button } from "@fluentui/react-northstar";
import { useGraph } from "./lib/useGraph";
import { ProfileCard } from "./ProfileCard";

export function Graph() {
  const { loading, error, data, reload } = useGraph(
    async (graph) => {
      const profile = await graph.api("/me").get();
      // exchange later, for a list of one profiles instead
      // const oneProfiles = await graph.api("/users").get(); // Mitarbeiter sind in der values liste (value: {me, you, theOtherOne, ...})
      let photoUrl = "";
      var availability = "Teams Verfügbarkeit: unbekannt"; // Output "unbekannt" = query hat nicht geklappt. Zugriffsberechtigungen muessen eingestellt sein.
      try {
        const photo = await graph.api("/me/photo/$value").get();
        photoUrl = URL.createObjectURL(photo);
        // availability = await graph.api("/beta/users/me/presence").get(); // BETA FUNCTIONALITY!
        availability = await graph.api("/me/presence").version('beta').get();
        document.write("DANIEL: ",availability);
      } catch {
        // Could not fetch photo from user's profile, return empty string as placeholder.
      }
      return { profile, photoUrl, availability };
    },
    { scope: ["User.Read"] }
  );

  return (
    <div>
      <h2>Get the user's profile photo</h2>
      <p>Click below to authorize this app to read your profile photo using Microsoft Graph.</p>
      <Button primary content="Authorize" disabled={loading} onClick={reload} />
      {loading && ProfileCard(true)}
      {!loading && error && (
        <div className="error">
          Failed to read your profile. Please try again later. <br /> Details: {error.toString()}
        </div>
      )}
      {!loading && data && ProfileCard(false, data)}
    </div>
  );
}
