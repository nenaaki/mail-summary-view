import { Fragment } from "react";

function isShortStringArray(value) {
  return (
    Array.isArray(value) &&
    value.length > 0 &&
    value.every((v) => typeof v === "string" && v.length <= 30)
  );
}

function FieldValue({ value, ancestors }) {
  if (value === null || value === undefined) {
    return <span className="dv-null">-</span>;
  }

  if (Array.isArray(value)) {
    if (value.length === 0) return <span className="dv-null">-</span>;

    if (isShortStringArray(value)) {
      return (
        <div className="dv-tags">
          {value.map((v, i) => (
            <span key={i} className="dv-tag">{v}</span>
          ))}
        </div>
      );
    }

    if (value.some((item) => item !== null && typeof item === "object" && !Array.isArray(item))) {
      return <ArrayContainer items={value} ancestors={ancestors} />;
    }

    return <span>{value.map(String).join(", ")}</span>;
  }

  if (typeof value === "object") {
    return <ObjectGrid obj={value} ancestors={ancestors} />;
  }

  return <span>{String(value)}</span>;
}

export function ObjectGrid({ obj, ancestors = new Set() }) {
  if (ancestors.has(obj)) {
    return <span className="dv-null">-</span>;
  }

  const next = new Set(ancestors);
  next.add(obj);

  return (
    <div className="dv-grid">
      {Object.entries(obj).map(([key, value]) => (
        <Fragment key={key}>
          <div className="dv-label">{key}</div>
          <div className="dv-value">
            <FieldValue value={value} ancestors={next} />
          </div>
        </Fragment>
      ))}
    </div>
  );
}

export function ArrayContainer({ items, ancestors = new Set() }) {
  return (
    <div className="dv-array">
      {items.map((item, i) => (
        <div key={i} className="dv-row">
          {item === null || item === undefined ? (
            <span className="dv-null">-</span>
          ) : typeof item === "object" && !Array.isArray(item) ? (
            <ObjectGrid obj={item} ancestors={ancestors} />
          ) : (
            <span>{String(item)}</span>
          )}
        </div>
      ))}
    </div>
  );
}
