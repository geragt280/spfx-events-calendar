import * as React from 'react';
import { FontIcon, Text } from 'office-ui-fabric-react'; // Assuming these components are available
import { IEventData } from './Interfaces/IEventData';
import * as moment from 'moment';

interface TooltipProps {
  headerColor: string;
  selectedEvent: IEventData;
  handleCloseTooltip: () => void;
}

const Tooltip: React.FC<TooltipProps> = (props) => {
  const tooltipStyle: React.CSSProperties = {
    position: "absolute",
    top: 195,
    left: 200,
    backgroundColor: props.headerColor,
    minHeight: 100,
    zIndex: 1000,
    padding: 20,
    maxWidth: 300,
    minWidth: 150,
    boxShadow: "rgba(0, 0, 0, 0.3) 4px 2px 5px"
  };

  
  // convert start and end date into moments so that we can manipulate them
  const startMoment: moment.Moment = moment(props.selectedEvent.EventDate);

  // event actually ends one second before the end date
  const endMoment: moment.Moment = moment(props.selectedEvent.EndDate).add(-1, "s");

  return (
    <div className="tooltipa" style={tooltipStyle}>
      <div>
        <FontIcon
          iconName="Cancel"
          style={{ float: "right", cursor: 'pointer' }}
          onClick={props.handleCloseTooltip}
        />
      </div>
      <div>
        <Text style={{ fontWeight: "bold" }}>
          {props.selectedEvent.title}
        </Text>
      </div>
      <div>
        <p dangerouslySetInnerHTML={{ __html: props.selectedEvent.Description }} />
      </div>
      <div>
        <Text><b>Start Date:</b> {startMoment.format("dddd, MMMM Do YYYY")}</Text>
      </div>
      <div>
        <Text><b>End Date:</b> {endMoment.format("dddd, MMMM Do YYYY")}</Text>
      </div>
    </div>
  );
};

export default Tooltip;
