# npoi-load-test

 ab -n 100 -c 100 http://localhost:59701/DownloadExcel.aspx                        
This is ApacheBench, Version 2.3 <$Revision: 1807734 $>                             
Copyright 1996 Adam Twiss, Zeus Technology Ltd, http://www.zeustech.net/            
Licensed to The Apache Software Foundation, http://www.apache.org/                  
                                                                                    
Benchmarking localhost (be patient).....done                                        
                                                                                    
                                                                                    
Server Software:        Microsoft-IIS/10.0                                          
Server Hostname:        localhost                                                   
Server Port:            59701                                                       
                                                                                    
Document Path:          /DownloadExcel.aspx                                         
Document Length:        666100 bytes                                                
                                                                                    
Concurrency Level:      100                                                         
Time taken for tests:   92.238 seconds                                              
Complete requests:      100                                                         
Failed requests:        83                                                          
   (Connect: 0, Receive: 0, Length: 83, Exceptions: 0)                              
Total transferred:      66653211 bytes                                              
HTML transferred:       66609911 bytes                                              
Requests per second:    1.08 [#/sec] (mean)                                         
Time per request:       92238.284 [ms] (mean)                                       
Time per request:       922.383 [ms] (mean, across all concurrent requests)         
Transfer rate:          705.68 [Kbytes/sec] received                                
                                                                                    
Connection Times (ms)                                                               
              min  mean[+/-sd] median   max                                         
Connect:        0    0   0.3      0       1                                         
Processing:  3502 48153 26820.5  47821   92220                                      
Waiting:     3476 48047 26804.9  47796   92205                                      
Total:       3502 48153 26820.5  47821   92221                                      
                                                                                    
Percentage of the requests served within a certain time (ms)                        
  50%  47821                                                                        
  66%  62520                                                                        
  75%  73597                                                                        
  80%  75853                                                                        
  90%  85732                                                                        
  95%  90542                                                                        
  98%  91957                                                                        
  99%  92221                                                                        
 100%  92221 (longest request)       

 excel在本机做了个简单的写测试，
100个并发用户，生成一个excel文件包含100个sheet（学生），每个sheet写100行20列，响应结果如上