function StarIcon({ active, size }) {
  const fill = active ? "#fbbf24" : "transparent";
  const stroke = active ? "#f59e0b" : "#cbd5e1";

  return (
    <svg width={size} height={size} viewBox="0 0 24 24" aria-hidden="true">
      <path
        d="M12 2.8 14.9 8.7l6.5.9-4.7 4.6 1.1 6.5L12 17.8 6.2 20.7l1.1-6.5L2.6 9.6l6.5-.9L12 2.8Z"
        fill={fill}
        stroke={stroke}
        strokeWidth="1.8"
        strokeLinejoin="round"
      />
    </svg>
  );
}

export default function StarRating({ rating, setRating, editable = false, size = 18 }) {
  return (
    <div className="star-row">
      {[1, 2, 3, 4, 5].map((value) => {
        const active = value <= rating;
        return (
          <button
            key={value}
            type="button"
            onClick={editable ? () => setRating(rating === value ? 0 : value) : undefined}
            disabled={!editable}
            className="star-button"
          >
            <StarIcon active={active} size={size} />
          </button>
        );
      })}
      {editable ? (
        <button type="button" className="mini-clear" onClick={() => setRating(0)}>
          Set 0
        </button>
      ) : null}
      <span className="star-value">{rating}/5</span>
    </div>
  );
}
