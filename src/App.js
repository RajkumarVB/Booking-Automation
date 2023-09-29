import './App.scss';
import {useState, useEffect} from "react";
import SideBar from './components/SideBar';
import { TATKAL, TATKAL_FOOD, CREDIT_CARD, SEVA, ACCOMODATION, SEVA_LOGIN,RS_300} from './utils/urls';

function App() {
  const [data,setData] = useState({});
  const [navBarItem,setNavBarItem] = useState("Tatkal");
  const [url,setUrl] = useState(TATKAL);

  useEffect(()=>{
    fetchData();
  },[url])

  const handleClick = (event) => {
    if(event.target.innerText === "Tatkal"){
      setNavBarItem("Tatkal");
      setUrl(TATKAL);
    } else if(event.target.innerText === "Tatkal Food"){
      setNavBarItem("Tatkal Food");
      setUrl(TATKAL_FOOD);
    } else if(event.target.innerText === "Credit Card"){
      setNavBarItem("Credit Card");
      setUrl(CREDIT_CARD);
    } else if(event.target.innerText === "Seva"){
      setNavBarItem("Seva");
      setUrl(SEVA);
    } else if(event.target.innerText === "Accomodation"){
      setNavBarItem("Accomodation");
      setUrl(ACCOMODATION);
    }  else if(event.target.innerText === "Seva Login"){
      setNavBarItem("Seva Login");
      setUrl(SEVA_LOGIN);
    } else if(event.target.innerText === "300"){
      setNavBarItem("300");
      setUrl(RS_300);
    }
  }
  const fetchData = async () => {
    const data = await fetch(url);
    const json = await data.json(); 
    setData(json);
    // setTatkalData(json);
  }

  return (
    <>
    <h1 id='title'>Booking Automation!🙏</h1>
    <ul className='nav-bar'>
      <li><button className={navBarItem === "Tatkal" ? "button active" : "button"} onClick={handleClick} type='button'>Tatkal</button></li>
      <li><button className={navBarItem === "Tatkal Food" ? "button active" : "button"}onClick={handleClick} type='button'>Tatkal Food</button></li>
      <li><button className={navBarItem === "300" ? "button active" : "button"}onClick={handleClick} type='button'>300</button></li>
      <li><button className={navBarItem === "Credit Card" ? "button active" : "button"}onClick={handleClick} type='button'>Credit Card</button></li>
      <li><button className={navBarItem === "Seva" ? "button active" : "button"}onClick={handleClick} type='button'>Seva</button></li>
      <li><button className={navBarItem === "Accomodation" ? "button active" : "button"}onClick={handleClick} type='button'>Accomodation</button></li>
      <li><button className={navBarItem === "Seva Login" ? "button active" : "button"}onClick={handleClick} type='button'>Seva Login</button></li>
    </ul>
    <SideBar data={data} menu={navBarItem}/>
    </>
  );
}

export default App;
