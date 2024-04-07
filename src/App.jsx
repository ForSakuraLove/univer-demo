import UniverSheet from './components/UniverSheet';
// import MyPieChart from './components/EChart'

function App() {
 
  return (
    <div id="root">
      <div style={{ display: 'flex', flexDirection: 'column', height: '100%' }}>
        <UniverSheet style={{ flex: 1 }}/>
        {/* <MyPieChart /> */}
      </div>
    </div>
  );
}

export default App;
