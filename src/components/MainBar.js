import { useEffect, useState } from "react";
const MainBar = (props) => {
    // let data = props.data;
    // let isDataArray = props.isDataArray; 
    const [data, setData] = useState([]);
    const [isDataArray, setIsDataArray] = useState(false);

    useEffect(() => {
        setData(props.data);
        setIsDataArray(props.isDataArray);
    }, [props.data, props.isDataArray]);
    return (
        <div>
            {isDataArray && data ? (
                data.map(person => {
                    return (
                        <div className="individual">
                            {Object.entries(person).map(entry => {
                                return (
                                    <p>{entry[0]} : {entry[1]}</p>
                                );
                            })}
                        </div>
                    );
                })
            ) : (
                <div className="one-person-container" key={data}>
                    <p>{data}</p>
                </div>
            )}
        </div>
    );
}

export default MainBar; 