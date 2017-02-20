"use strict";

['it', 'xit', 'fit'].forEach(method => {
    global[`${method}Async`] = (message, spec) => {
        return global[method](message, done => {
            spec().then(done).catch(done.fail);
        });
    };
});
