import { useEffect, useRef, useState } from "react";
import { REGIONS } from "./directoryData";
import { getRegionStrength } from "./directoryUtils";

function styleMap(container, installersByRegion, selectedRegion, hoveredRegion) {
  if (!container) return;
  const maxCount = Math.max(...REGIONS.map((region) => installersByRegion[region.id]?.length || 0), 0);

  REGIONS.forEach((region) => {
    const group = container.querySelector(`#${CSS.escape(region.svgId)}`);
    const shape = group?.querySelector("path.st1");
    if (!shape) return;

    let fill = getRegionStrength(installersByRegion[region.id]?.length || 0, maxCount);
    if (hoveredRegion === region.id) fill = "#75508c";
    if (selectedRegion === region.id) fill = "#5f3c74";

    shape.style.fill = fill;
    shape.style.cursor = "pointer";
    shape.style.transition = "fill 160ms ease";
    shape.style.stroke = "#ffffff";
    shape.style.strokeWidth = "3px";
  });
}

export default function InteractiveMap({
  installersByRegion,
  selectedRegion,
  setSelectedRegion,
  hoveredRegion,
  setHoveredRegion
}) {
  const mapRef = useRef(null);
  const [svgMarkup, setSvgMarkup] = useState("");

  useEffect(() => {
    let cancelled = false;

    fetch("/uk-map.svg")
      .then((response) => response.text())
      .then((text) => {
        if (!cancelled) setSvgMarkup(text);
      })
      .catch((error) => console.error("Could not load map SVG", error));

    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    styleMap(mapRef.current, installersByRegion, selectedRegion, hoveredRegion);
  }, [hoveredRegion, installersByRegion, selectedRegion, svgMarkup]);

  useEffect(() => {
    const container = mapRef.current;
    if (!container || !svgMarkup) return undefined;

    const cleanups = [];
    REGIONS.forEach((region) => {
      const group = container.querySelector(`#${CSS.escape(region.svgId)}`);
      const shape = group?.querySelector("path.st1");
      if (!shape) return;

      const onEnter = () => setHoveredRegion(region.id);
      const onLeave = () => setHoveredRegion(null);
      const onClick = () => setSelectedRegion(region.id === selectedRegion ? null : region.id);

      shape.addEventListener("mouseenter", onEnter);
      shape.addEventListener("mouseleave", onLeave);
      shape.addEventListener("click", onClick);

      cleanups.push(() => {
        shape.removeEventListener("mouseenter", onEnter);
        shape.removeEventListener("mouseleave", onLeave);
        shape.removeEventListener("click", onClick);
      });
    });

    return () => cleanups.forEach((cleanup) => cleanup());
  }, [selectedRegion, setHoveredRegion, setSelectedRegion, svgMarkup]);

  return (
    <div className="map-shell">
      <img className="map-shell-logo" src="/branding/signs-express-logo.svg" alt="Signs Express" />
      <div ref={mapRef} className="map-markup" dangerouslySetInnerHTML={{ __html: svgMarkup }} />
    </div>
  );
}
