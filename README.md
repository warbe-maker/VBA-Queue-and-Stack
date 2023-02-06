## Common-VBA-Queue-and-Stack-services
Common VBA Components providing Queue (mQueue, clsQueue) and Stack (mStack, clsStack) services.
## Services
All services below are provided either by a Standard-Module (mQueue, mStack) by a Class-Module (clsQueue, clsStack). In addition the services may be integrated as Private procedures (Queue...., Stack....) which are identical in the clsQueue/mQueue and the clsStack/mStack modules.

| Service   | Queue | Stack | Description                                                                         |
| --------- |:-----:|:-----:|---------------------------------------------------------------------- |
| Bottom    |       |   x   | Returns the bottom item on the stack                                                |
| DeQueue   |       |       | De-queues (returns and removes):<br>- the first item in the queue (the default)<br>- a specific item plus its position provided it is unique in the queue<br>- an item identified by its position    |
| EnQueue   |   x   |       | Adds/en-queues an item
| First     |   x   |       | Returns the first item added/en-queued)  |
| IsEmpty   |   x   |   x   | Returns TRUE when the stack is empty                                                |
| IsStacked |       |   x   | Returns TRUE and its position when a provided item is on the stack                  |
| Item      |   x   |   x   | Returns an item on a provided position                                              |
| Last      |   x   |       | Returns the last item in the queue (i.e. the last item added/en-queued)
| Pop       |       |   x   | Returns the top (last added) item from the stack and removes it                    |
| Push      |       |   x   | Pushes an item on the stack                                                        |
| Size      |   x   |   x   | Returns the current stack's size                                                   |
| Top       |       |   x   | Returns the top item on the stack by leaving it on the stack, i.e. not popping it. |


## Installation and usage options
### The services as ClassModules
By far the best and most elegant choice. Once installed the services are available for as many queues/stacks as required and the services are prefixed with the queue- or stack- name in accordance with the applications means.

### The services as StandardModules
### The services Queue..../Stack..... as private procedures
The procedures Queue.... (from within the _clsQueue_ or the _mQueue_ modules may be copied directly into any StandardModule. This option may be applicable when a module's implementation aims for independence of any other module. Using the ClassModule remains the best choice otherwise. 





