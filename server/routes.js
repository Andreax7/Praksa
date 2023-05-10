
const express = require("express");

const routes = express.Router();

const fs = require('fs');
const data = fs.readFileSync('data.json', 'utf8');


routes.get('/user/:id', async (request, response) => {
    try{ 
        const jsonData = JSON.parse(data);
        const id = request.params.id
        let usersArr = jsonData.users
        for( let u in usersArr){
            if(usersArr[u].id == id){
                return response.status(200).json(usersArr[u]); 
            }
        }
        return response.status(500).send("no such user");
    }
    catch (error) {
     return response.status(500).send(error);
    }
    
  });

  routes.get('/post/:id', async (request, response) => {
    try{ 
        const jsonData = JSON.parse(data);
        const id = request.params.id
        let postsArr = jsonData.posts
        for( let p in postsArr){
            if(postsArr[p].id == id){
                return response.status(200).json(postsArr[p]); 
            }
        }
        return response.status(500).send("no such post");
    }
    catch (error) {
     return response.status(500).send(error);
    }
    
  });


  routes.get('/postByDate', async (request, response) => {
    try{ 
        const jsonData = JSON.parse(data); // salje se u obliku { from:"yyyy-mm-dd", to:"yyyy-mm-dd"}
        const from = request.body.from
        const to = request.body.to
        var postsLimit = []
        let postsArr = jsonData.posts
        for( let p in postsArr){
            var datumPart = postsArr[p].last_update
            const datum = datumPart.substr(0,10)
            
            console.log("matching yr", request.body.to,  to.substr(0,4), datum.substr(0,4) > from.substr(0,4), datum.substr(0,4) >= to.substr(0,4) )
            if(datum.substr(0,4) >= from.substr(0,4) || datum.substr(0,4) >= to.substr(0,4)) // if year match
            {
                if(datum.substr(5,7) >= from.substr(5,7) || datum.substr(5,7) >= to.substr(5,7)) // if month match
                {
                    if(datum.substr(8,10) >= from.substr(8,10) && datum.substr(8,10) >= to.substr(8,10)) // if day match
                    {
                        console.log("matching day", datum.substr(8,10) > from.substr(8,10), datum.substr(8,10) >= to.substr(8,10) )
                        postsLimit.push(postsArr[p])

                    }
                }
            }
    
        }
        return response.status(200).send(JSON.stringify(postsLimit));
    }
    catch (error) {
     return response.status(500).send(error);
    }
    
  });


  module.exports = routes;