import * as React from 'react';
import styles from './Spotlight.module.scss';
import { SPFI } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';

interface Employee {
    employeeName: string;
    employeePhoto: string; // Assuming this is a string representation (URL or base64)
    eventBlessing: string;
    eventDescription: string;
    isDisplay: boolean;
    Index: number;
    ID: number;
}

interface SpotlightProps {
    sp: SPFI;
    SpotlightId: string;
    context: WebPartContext;
    SpotlightInterval: number;
}

interface SpotlightState {
    currentIndex: number;
    employeesList: Employee[];
    isLoading: boolean;
}

const defaultSpotlightState: SpotlightState = {
    currentIndex: 0,
    employeesList: [],
    isLoading: true, // Introduce a loading state
};

export default function Spotlight({ sp, SpotlightId, context, SpotlightInterval }: SpotlightProps) {
    const [state, setState] = React.useState<SpotlightState>(defaultSpotlightState);

    React.useEffect(() => {
        const fetchSpotlight = async () => {
            const items: Employee[] = await sp.web.lists.getById(SpotlightId).items
                .select('employeeName', 'employeePhoto', 'eventBlessing', 'eventDescription', 'isDisplay', 'Index', 'ID')
                .orderBy('Index')()
                .then(items => items.filter(item => item.isDisplay)); // Assuming you want to filter by `isDisplay`
            setState(prevState => ({
                ...prevState, employeesList: items,
                isLoading: false // Update loading state upon completion
            }));
        };

        fetchSpotlight();
    }, []);

    React.useEffect(() => {
        const interval = setInterval(() => {
            setState(prevState => ({
                ...prevState,
                currentIndex: (prevState.currentIndex + 1) % prevState.employeesList.length,
            }));
        }, SpotlightInterval);

        return () => clearInterval(interval); // Cleanup the interval on component unmount
    }, [state.isLoading, state.employeesList.length]);

    const currentEmployee = state?.employeesList[state.currentIndex];

    const onPictureConverterUrl = (imageFromSP: string, id: number): string => {
        let imageUrl
        if (JSON.parse(imageFromSP)?.fileName) {

            imageUrl = context.pageContext.web.absoluteUrl + "/Lists/EmployeeSpotlight/Attachments/" + id + "/" + JSON.parse(imageFromSP)?.fileName;
        } else {
            imageUrl = require("../../assets/profilePicDemo.jpg")
        }

        return imageUrl;
    };

    return (
        <div className={styles.spotlightContainer}>
            <span className={`${styles.title}`}>
                זרקור לעובד
            </span>
            {currentEmployee && (
                <div className={`${styles.item}`}>
                    <div className={styles.imageContainer}>
                        <img
                            src={onPictureConverterUrl(currentEmployee?.employeePhoto, currentEmployee?.ID)}
                            alt={currentEmployee?.employeeName}
                            className={styles.peopleImg}
                        />
                    </div>
                    <div dangerouslySetInnerHTML={{ __html: currentEmployee?.eventDescription }} className={styles.eventDescription}
                        style={{ overflowY: currentEmployee?.eventDescription.length < 150 ? 'hidden' : 'scroll' }}></div>
                    <span>{currentEmployee?.employeeName}</span>
                </div>
            )}
        </div>
    );
}
