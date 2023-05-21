
const express = require("express")
const routes = express.Router()

const fs = require('fs')
const data = fs.readFileSync('data.json', 'utf8')


//*****    GET  *******/
routes.get('/user/:id', async (request, response) => {
    try{ 
        const jsonData = JSON.parse(data)
        const id = request.params.id
        let usersArr = jsonData.users
        for( let u in usersArr){
            if(usersArr[u].id == id){
                return response.status(200).json(usersArr[u]) 
            }
        }
        return response.status(500).send("no such user")
    }
    catch (error) {
        return response.status(500).send(error)
    }
    
  });

  routes.get('/post/:id', async (request, response) => {
    try{ 
        const jsonData = JSON.parse(data)
        const id = request.params.id
        let postsArr = jsonData.posts

        for( let p in postsArr){
            if(postsArr[p].id == id){
                return response.status(200).json(postsArr[p]) 
            }
        }
        return response.status(500).send("no such post")
    }
    catch (error) {
        return response.status(500).send(error)
    }
    
  });


  routes.get('/postByDate', async (request, response) => {
    try{ 
        const jsonData = JSON.parse(data) // salje se u obliku { from:"yyyy-mm-dd", to:"yyyy-mm-dd"}
        const from = request.body.from
        const to = request.body.to
        var postsLimit = []
        let postsArr = jsonData.posts

            for( let p in postsArr){
                var datumPart = postsArr[p].last_update
                const datum = datumPart.substr(0,10)  
                //console.log("matching yr", request.body.to,  to.substr(0,4), datum.substr(0,4) > from.substr(0,4), datum.substr(0,4) >= to.substr(0,4) )
                if(datum.substr(0,4) >= from.substr(0,4) || datum.substr(0,4) >= to.substr(0,4)){ // if year match
    
                    if(datum.substr(5,7) >= from.substr(5,7) || datum.substr(5,7) >= to.substr(5,7)){ // if month match
                        
                        if(datum.substr(8,10) >= from.substr(8,10) && datum.substr(8,10) >= to.substr(8,10)){ // if day match
                            //console.log("matching day", datum.substr(8,10) > from.substr(8,10), datum.substr(8,10) >= to.substr(8,10) )
                            postsLimit.push(postsArr[p])
                        }
                    }
                }
            }
        return response.status(200).send(JSON.stringify(postsLimit))
    }
    catch (error) {
        return response.status(500).send(error)
    }
  });

  //*****    POST  *******/

  routes.post('/user/:uid',async (request, response) => {
    try{  
        let jsonData = JSON.parse(data)
        let usersArr = jsonData.users
        const user_id = request.params.uid
        const email = request.body.newEmail
        
        if(!userCheck(user_id) || email === undefined){
            response.status(500).send('user does not exist or empty email')
            return new Error("user does not exist")
        }
        else{
            for( let u in usersArr ){
                if(usersArr[u].id == parseInt(user_id)){
                    usersArr[u].email = email
                    fs.writeFileSync('data.json', JSON.stringify(jsonData)); 
                    return response.status(200).json(usersArr[u]) 
                }
            }
        }
    }  
        catch(error) {
          response.status(500).send(error)
        }
  });

  //*****    PUT  *******/

    routes.put('/user/:id/newpost/',async (request, response) => {
        try {       
            //check if user exists
            const user_id = request.params.id
            
            if(!userCheck(user_id)){
                response.status(500).send('user does not exist')
                return new Error("user does not exist")
            }
            else{
                let jsonData = JSON.parse(data)
                let postsArr = jsonData.posts

                const title = request.body.title
                const body = request.body.body
                const last_update = stringFormatDate(new Date().toLocaleString())
        
                const id = parseInt((postsArr.length)+1)
                const newPost = {
                    "id":id, 
                    "title":title, 
                    "body":body,
                    "user_id": parseInt(user_id), 
                    "last_update": last_update
                }
                postsArr.push(newPost)
                //console.log(postsArr)
                fs.writeFileSync('data.json', JSON.stringify(jsonData));  
                //console.log(newPost)  
                return response.status(200).send(JSON.stringify(newPost))     
            }       
        }  
        catch(error) {
            response.status(500).send(error)
        }
    });

  //******************************************
  //***********HELPER FUNCTIONS **************


    function stringFormatDate(dateStr){
        var newDateStr = dateStr[8]+dateStr[9]+dateStr[10]+dateStr[11]+'-'+ dateStr[4]+dateStr[5] +'-'+ dateStr[0]+dateStr[1]+' ' + dateStr[14]+ dateStr[15]+ dateStr[16]+ dateStr[17]+ dateStr[18]+ dateStr[19]+dateStr[20]+dateStr[21]
        return newDateStr
    }

    function userCheck(userId){
        const jsonData = JSON.parse(data)
        let usersArr = jsonData.users
    
            for( let u in usersArr){
                if(usersArr[u].id == parseInt(userId)){
                    return true 
                }
            }
        return false
    }


  module.exports = routes