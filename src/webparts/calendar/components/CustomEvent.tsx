import * as React from 'react';
import { IEventData } from './Interfaces/IEventData';

const CustomEvent: React.FC<{event:IEventData}> = ({ event }) => {
  const [showTooltip, setShowTooltip] = React.useState(false);

  const handleMouseEnter = () => {
    console.log("tooltip opened");
    setShowTooltip(true);
  };

  const handleMouseLeave = () => {
    console.log("tooltip closed", event.Description);
    setShowTooltip(false);
  };

  return (
    <div
      onMouseEnter={handleMouseEnter}
      onMouseLeave={handleMouseLeave}
      style={{ position: 'relative', cursor: 'pointer'  }}
    >
      {event.title}
      {showTooltip && (
        <div
          style={{
            // position: 'absolute',
            top: '100%',
            left: 0,
            background: 'rgba(255, 255, 0, 0.8)',
            padding: '5px',
            border: '1px solid #ccc',
            borderRadius: '5px',
            color:'black',
            zIndex: 1000,
          }}
        >
          {event.title + " " + event.Description || 'Default Tooltip Content'}
        </div>
      )}
    </div>
  );
};

export default CustomEvent;
